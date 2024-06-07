package com.example.MachinationsServer.Service;

import com.example.MachinationsServer.Models.Diagram;
import com.example.MachinationsServer.Models.Node;
import com.example.MachinationsServer.Repository.*;
import jakarta.persistence.EntityNotFoundException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.NodeList;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

@Service
public class DiagramService {

    @Autowired
    private IDiagramRepo diagramRepo;

    @Autowired
    private INodeRepo nodeRepo;

    public void saveDiagram(String name, ArrayList<Node> nodes) {
        Optional<Diagram> databaseDiagramOpt = diagramRepo.findByName(name);
        Diagram diagram;
        if (databaseDiagramOpt.isPresent()) {
            diagram = databaseDiagramOpt.get();
            nodeRepo.deleteAll(diagram.getNode_list());
            diagram.getNode_list().clear();
        } else {
            diagram = Diagram.builder()
                    .name(name)
                    .node_list(new ArrayList<>())
                    .build();
        }
        for (Node node : nodes) {
            node.setDiagram(diagram);
            diagram.getNode_list().add(node);
        }
        diagramRepo.save(diagram);
        nodeRepo.saveAll(diagram.getNode_list());
    }


    public List<Diagram> getAllDiagrams(){
        return diagramRepo.findAll();
    }

    public void saveToXLSX(String fileName, ArrayList<Node> nodes) throws IOException{
        Path path = Paths.get(System.getProperty("user.home"), "Desktop", fileName + ".xlsx");
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream(path.toString())) {
            Sheet sheet = workbook.createSheet("Data");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);

            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            CellStyle columnCellStyle = workbook.createCellStyle();
            columnCellStyle.setFont(headerFont);

            Font columnFont = workbook.createFont();
            columnFont.setBold(true);
            headerCellStyle.setFont(columnFont);

            columnCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            columnCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            columnCellStyle.setBorderTop(BorderStyle.THIN);
            columnCellStyle.setBorderBottom(BorderStyle.THIN);

            LocalDateTime now = LocalDateTime.now();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss");
            String formattedDateTime = now.format(formatter);

            String[] attributes = {
                    "","",
                    "Diagram Properties", "",
                    "", "",
                    "Name:", fileName,
                    "Url:", "",
                    "Owner","Calotă Ovidiu",
                    "Creation Date:", "",
                    "Last Change:", formattedDateTime,
                    "Time Interval:", "1",
                    "Time Steps Limit:", "100",
                    "Number of Runs (total):", "20",
                    "Exporting Codec:", "v2.0",

            };

            int rowIndex = 0;
            for (int i = 0; i < attributes.length; i += 2) {
                Row row = sheet.createRow(rowIndex++);
                if (!attributes[i].isEmpty()) {
                    if(attributes[i].equals("Diagram Properties"))
                    {
                        Cell headerCell = row.createCell(0);  // Start from second column for header
                        headerCell.setCellValue(attributes[i]);
                        headerCell.setCellStyle(headerCellStyle);
                    }
                    else{
                        Cell headerCell = row.createCell(1);  // Start from second column for header
                        headerCell.setCellValue(attributes[i]);
                        headerCell.setCellStyle(headerCellStyle);
                    }
                    Cell valueCell = row.createCell(2);  // Value in the third column
                    valueCell.setCellValue(attributes[i + 1]);
                }
            }

            Map<String ,ArrayList<Node>> splitNodes = splitByType(nodes);

            Row row = sheet.createRow(rowIndex++);
            if(splitNodes.containsKey("Source")) {
                if (!splitNodes.get("Source").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Sources");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Activation");

                    cell = columns.createCell(7);
                    cell.setCellValue("Resources (color)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Activation Mode");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(9);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);

                }
                for (Node source : splitNodes.get("Source")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(source.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(source.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + source.getX() + ",\"y\":" + source.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=source-shape;whiteSpace=wrap;html=1;strokeWidth=" + source .getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + source.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    switch (source.getActivationMode()) {
                        case "2" -> cell.setCellValue("interactive");
                        case "3" -> cell.setCellValue("automatic");
                        case "4" -> cell.setCellValue("onstart");
                        default -> cell.setCellValue("passive");
                    }

                    cell = row.createCell(7);
                    cell.setCellValue("black");

                    cell = row.createCell(8);
                    cell.setCellValue("push-any");

                    cell = row.createCell(9);
                    cell.setCellValue(nodes.indexOf(source));
                }
            }

            if(splitNodes.containsKey("Drain")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("Drain").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Drains");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Activation");

                    cell = columns.createCell(7);
                    cell.setCellValue("Activation Mode");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node drain : splitNodes.get("Drain")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(drain.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(drain.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + drain.getX() + ",\"y\":" + drain.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=drain-shape;whiteSpace=wrap;html=1;strokeWidth=" + drain.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + drain.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    switch (drain.getActivationMode()) {
                        case "2" -> cell.setCellValue("interactive");
                        case "3" -> cell.setCellValue("automatic");
                        case "4" -> cell.setCellValue("onstart");
                        default -> cell.setCellValue("passive");
                    }

                    cell = row.createCell(7);
                    cell.setCellValue("pull-any");

                    cell = row.createCell(8);
                    cell.setCellValue(nodes.indexOf(drain));
                }
            }

            if(splitNodes.containsKey("Pool")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("Pool").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Pools");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Activation");

                    cell = columns.createCell(7);
                    cell.setCellValue("Activation Mode");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Resources");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(9);
                    cell.setCellValue("Resources (color)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(10);
                    cell.setCellValue("Capacity (limit)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(11);
                    cell.setCellValue("Capacity (display)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(12);
                    cell.setCellValue("Overflow");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(13);
                    cell.setCellValue("Show in chart");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(14);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);

                }
                for (Node pool : splitNodes.get("Pool")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(pool.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(pool.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + pool.getX() + ",\"y\":" + pool.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=pool-shape;whiteSpace=wrap;html=1;strokeWidth=" + pool.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + pool.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    switch (pool.getActivationMode()) {
                        case "2" -> cell.setCellValue("interactive");
                        case "3" -> cell.setCellValue("automatic");
                        case "4" -> cell.setCellValue("onstart");
                        default -> cell.setCellValue("passive");
                    }

                    cell = row.createCell(7);
                    cell.setCellValue("pull-any");

                    cell = row.createCell(8);
                    cell.setCellValue(0);

                    cell = row.createCell(9);
                    cell.setCellValue("Black");

                    cell = row.createCell(10);
                    cell.setCellValue(-1);

                    cell = row.createCell(11);
                    cell.setCellValue(25);

                    cell = row.createCell(12);
                    cell.setCellValue("block");

                    cell = row.createCell(13);
                    cell.setCellValue(true);

                    cell = row.createCell(14);
                    cell.setCellValue(nodes.indexOf(pool));
                }
            }

            if(splitNodes.containsKey("Gate")){
                    sheet.createRow(rowIndex++);

                    row = sheet.createRow(rowIndex++);
                    if (!splitNodes.get("Gate").isEmpty()) {
                        Cell cell = row.createCell(0);
                        cell.setCellValue("Gates");
                        cell.setCellStyle(headerCellStyle);

                        Row columns = sheet.createRow(rowIndex++);
                        cell = columns.createCell(0);
                        cell.setCellValue("ID");
                        cell.setCellStyle(columnCellStyle);

                        cell = columns.createCell(1);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Label");

                        cell = columns.createCell(2);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Layer ID");

                        cell = columns.createCell(3);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Group ID");

                        cell = columns.createCell(4);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Geometry");

                        cell = columns.createCell(5);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Style");

                        cell = columns.createCell(6);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Activation");

                        cell = columns.createCell(7);
                        cell.setCellValue("Activation Mode");
                        cell.setCellStyle(columnCellStyle);

                        cell = columns.createCell(8);
                        cell.setCellValue("Distribution");
                        cell.setCellStyle(columnCellStyle);

                        cell = columns.createCell(9);
                        cell.setCellValue("Position");
                        cell.setCellStyle(columnCellStyle);

                    }
                    for (Node gate : splitNodes.get("Gate")) {
                        row = sheet.createRow(rowIndex++);
                        Cell cell = row.createCell(0);
                        cell = row.createCell(0);
                        cell.setCellValue(gate.getNodeId());

                        cell = row.createCell(1);
                        cell.setCellValue(gate.getLabel());

                        cell = row.createCell(2);
                        cell.setCellValue(201);

                        cell = row.createCell(4);
                        cell.setCellValue("{\"x\":" + gate.getX() + ",\"y\":" + gate.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                        cell = row.createCell(5);
                        cell.setCellValue("shape=gate-shape;whiteSpace=wrap;html=1;strokeWidth=" + gate.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + gate.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                        cell = row.createCell(6);
                        switch (gate.getActivationMode()) {
                            case "2" -> cell.setCellValue("interactive");
                            case "3" -> cell.setCellValue("automatic");
                            case "4" -> cell.setCellValue("onstart");
                            default -> cell.setCellValue("passive");
                        }
                        cell = row.createCell(7);
                        cell.setCellValue("pull-any");

                        cell = row.createCell(8);
                        cell.setCellValue("deterministic");

                        cell = row.createCell(9);
                        cell.setCellValue(nodes.indexOf(gate));
                    }
                }

            if(splitNodes.containsKey("Trader")){
                    sheet.createRow(rowIndex++);

                    row = sheet.createRow(rowIndex++);
                    if (!splitNodes.get("Trader").isEmpty()) {
                        Cell cell = row.createCell(0);
                        cell.setCellValue("Traders");
                        cell.setCellStyle(headerCellStyle);

                        Row columns = sheet.createRow(rowIndex++);
                        cell = columns.createCell(0);
                        cell.setCellValue("ID");
                        cell.setCellStyle(columnCellStyle);

                        cell = columns.createCell(1);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Label");

                        cell = columns.createCell(2);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Layer ID");

                        cell = columns.createCell(3);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Group ID");

                        cell = columns.createCell(4);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Geometry");

                        cell = columns.createCell(5);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Style");

                        cell = columns.createCell(6);
                        cell.setCellStyle(columnCellStyle);
                        cell.setCellValue("Activation");

                        cell = columns.createCell(7);
                        cell.setCellValue("Trade");
                        cell.setCellStyle(columnCellStyle);

                        cell = columns.createCell(8);
                        cell.setCellValue("Position");
                        cell.setCellStyle(columnCellStyle);

                    }
                    for (Node trader : splitNodes.get("Trader")) {
                        row = sheet.createRow(rowIndex++);
                        Cell cell = row.createCell(0);
                        cell = row.createCell(0);
                        cell.setCellValue(trader.getNodeId());

                        cell = row.createCell(1);
                        cell.setCellValue(trader.getLabel());

                        cell = row.createCell(2);
                        cell.setCellValue(201);

                        cell = row.createCell(4);
                        cell.setCellValue("{\"x\":" + trader.getX() + ",\"y\":" + trader.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                        cell = row.createCell(5);
                        cell.setCellValue("shape=trader-shape;whiteSpace=wrap;html=1;strokeWidth=" + trader.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + trader.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                        cell = row.createCell(6);
                        switch (trader.getActivationMode()) {
                            case "2" -> cell.setCellValue("interactive");
                            case "3" -> cell.setCellValue("automatic");
                            case "4" -> cell.setCellValue("onstart");
                            default -> cell.setCellValue("passive");
                        }
                        cell = row.createCell(7);
                        cell.setCellValue("single");

                        cell = row.createCell(8);
                        cell.setCellValue(nodes.indexOf(trader));
                    }
                }

            if(splitNodes.containsKey("Converter")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("Converter").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Converters");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Activation");

                    cell = columns.createCell(7);
                    cell.setCellValue("Activation Mode");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Resources (color)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(9);
                    cell.setCellValue("Conversion");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(10);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node converter : splitNodes.get("Converter")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(converter.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(converter.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + converter.getX() + ",\"y\":" + converter.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=converter-shape;whiteSpace=wrap;html=1;strokeWidth=" + converter.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + converter.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    switch (converter.getActivationMode()) {
                        case "2" -> cell.setCellValue("interactive");
                        case "3" -> cell.setCellValue("automatic");
                        case "4" -> cell.setCellValue("onstart");
                        default -> cell.setCellValue("passive");
                    }
                    cell = row.createCell(7);
                    cell.setCellValue("pull-any");

                    cell = row.createCell(8);
                    cell.setCellValue("Black");

                    cell = row.createCell(9);
                    cell.setCellValue("single");

                    cell = row.createCell(10);
                    cell.setCellValue(nodes.indexOf(converter));
                }
            }

            if(splitNodes.containsKey("Register")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("Register").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Registers");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Formula");

                    cell = columns.createCell(7);
                    cell.setCellValue("Interactive");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Limit (minimum)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(9);
                    cell.setCellValue("Limit (maximum)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(10);
                    cell.setCellValue("Value (initial)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(11);
                    cell.setCellValue("Value (step)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(12);
                    cell.setCellValue("Show in chart");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(13);
                    cell.setCellValue("Interactive");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(14);
                    cell.setCellValue("Force update in each step");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(15);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node register : splitNodes.get("Register")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(register.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(register.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + register.getX() + ",\"y\":" + register.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=register-shape;whiteSpace=wrap;html=1;strokeWidth=" + register.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + register.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(7);
                    cell.setCellValue("false");

                    cell = row.createCell(10);
                    cell.setCellValue(0);

                    cell = row.createCell(11);
                    cell.setCellValue(1);

                    cell = row.createCell(12);
                    cell.setCellValue(true);

                    cell = row.createCell(13);
                    cell.setCellValue(false);

                    cell = row.createCell(13);
                    cell.setCellValue(false);

                    cell = row.createCell(15);
                    cell.setCellValue(nodes.indexOf(register));
                }
            }

            if(splitNodes.containsKey("Delay")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("Delay").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Delays");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Activation");

                    cell = columns.createCell(7);
                    cell.setCellValue("Queue");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node delay : splitNodes.get("Delay")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(delay.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(delay.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + delay.getX() + ",\"y\":" + delay.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=delay-shape;whiteSpace=wrap;html=1;strokeWidth=" + delay.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + delay.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    switch (delay.getActivationMode()) {
                        case "2" -> cell.setCellValue("interactive");
                        case "3" -> cell.setCellValue("automatic");
                        case "4" -> cell.setCellValue("onstart");
                        default -> cell.setCellValue("passive");
                    }
                    cell = row.createCell(7);
                    cell.setCellValue("false");

                    cell = row.createCell(8);
                    cell.setCellValue(nodes.indexOf(delay));
                }
            }

            if(splitNodes.containsKey("EndCondition")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("EndCondition").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("End Conditions");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node endCondition : splitNodes.get("EndCondition")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(endCondition.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(endCondition.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":" + endCondition.getX() + ",\"y\":" + endCondition.getY() + ",\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=end-condition-shape;whiteSpace=wrap;html=1;strokeWidth=" + endCondition.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + endCondition.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    cell.setCellValue(nodes.indexOf(endCondition));
                }
            }

            if(splitNodes.containsKey("ResourceConnection")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("ResourceConnection").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("Resource Connections");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellValue("Formula");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(7);
                    cell.setCellValue("Interval");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Source");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(9);
                    cell.setCellValue("Target");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(10);
                    cell.setCellValue("Transfer");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(11);
                    cell.setCellValue("Color Coding");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(12);
                    cell.setCellValue("Color Coding (color)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(13);
                    cell.setCellValue("Shuffle Source");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(14);
                    cell.setCellValue("Limits (minimum)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(15);
                    cell.setCellValue("Limits (maximum)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(16);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node resourceConnection : splitNodes.get("ResourceConnection")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(resourceConnection.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(resourceConnection.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":0,\"y\":0,\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=end-condition-shape;whiteSpace=wrap;html=1;strokeWidth=" + resourceConnection.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + resourceConnection.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(6);
                    cell.setCellValue(1);

                    cell = row.createCell(8);
                    cell.setCellValue(resourceConnection.getStartId());

                    cell = row.createCell(9);
                    cell.setCellValue(resourceConnection.getEndId());

                    cell = row.createCell(10);
                    cell.setCellValue("interval-based");

                    cell = row.createCell(11);
                    cell.setCellValue(false);

                    cell = row.createCell(12);
                    cell.setCellValue("Black");

                    cell = row.createCell(13);
                    cell.setCellValue(false);

                    cell = row.createCell(16);
                    cell.setCellValue(nodes.indexOf(resourceConnection));
                }
            }

            if(splitNodes.containsKey("StateConnection")) {
                sheet.createRow(rowIndex++);

                row = sheet.createRow(rowIndex++);
                if (!splitNodes.get("StateConnection").isEmpty()) {
                    Cell cell = row.createCell(0);
                    cell.setCellValue("State Connections");
                    cell.setCellStyle(headerCellStyle);

                    Row columns = sheet.createRow(rowIndex++);
                    cell = columns.createCell(0);
                    cell.setCellValue("ID");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(1);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Label");

                    cell = columns.createCell(2);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Layer ID");

                    cell = columns.createCell(3);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Group ID");

                    cell = columns.createCell(4);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Geometry");

                    cell = columns.createCell(5);
                    cell.setCellStyle(columnCellStyle);
                    cell.setCellValue("Style");

                    cell = columns.createCell(6);
                    cell.setCellValue("Formula");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(7);
                    cell.setCellValue("Source");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(8);
                    cell.setCellValue("Target");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(9);
                    cell.setCellValue("Color Coding");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(10);
                    cell.setCellValue("Color Coding (color)");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(11);
                    cell.setCellValue("Trigger on");
                    cell.setCellStyle(columnCellStyle);

                    cell = columns.createCell(12);
                    cell.setCellValue("Position");
                    cell.setCellStyle(columnCellStyle);
                }
                for (Node stateConnection : splitNodes.get("StateConnection")) {
                    row = sheet.createRow(rowIndex++);
                    Cell cell = row.createCell(0);
                    cell = row.createCell(0);
                    cell.setCellValue(stateConnection.getNodeId());

                    cell = row.createCell(1);
                    cell.setCellValue(stateConnection.getLabel());

                    cell = row.createCell(2);
                    cell.setCellValue(201);

                    cell = row.createCell(4);
                    cell.setCellValue("{\"x\":0,\"y\":0,\"width\":46,\"height\":46,\"offset\":{\"x\":0,\"y\":52},\"TRANSLATE_CONTROL_POINTS\":true,\"relative\":false,\"newFormat\":false,\"formulaOnTop\":true,\"vertexOrientation\":\"bottom\"}");

                    cell = row.createCell(5);
                    cell.setCellValue("shape=end-condition-shape;whiteSpace=wrap;html=1;strokeWidth=" + stateConnection.getSize() + ";aspect=fixed;resizable=0;fontSize=16;fontColor=#000000;strokeColor=" + stateConnection.getColor() + ";fillColor=#FFFFFF;verticalAlign=top;");

                    cell = row.createCell(7);
                    cell.setCellValue(stateConnection.getStartId());

                    cell = row.createCell(8);
                    cell.setCellValue(stateConnection.getEndId());

                    cell = row.createCell(9);
                    cell.setCellValue(false);

                    cell = row.createCell(10);
                    cell.setCellValue("Black");

                    cell = row.createCell(11);
                    cell.setCellValue("receiving resource");

                    cell = row.createCell(12);
                    cell.setCellValue(nodes.indexOf(stateConnection));
                }
            }

            sheet.createRow(rowIndex++);

            Row layers = sheet.createRow(rowIndex++);
            Cell cell = layers.createCell(0);
            cell.setCellValue("Layers");
            cell.setCellStyle(headerCellStyle);

            layers = sheet.createRow(rowIndex++);
            cell = layers.createCell(0);
            cell.setCellValue("ID");
            cell.setCellStyle(columnCellStyle);

            cell = layers.createCell(1);
            cell.setCellStyle(columnCellStyle);
            cell.setCellValue("Label");

            cell = layers.createCell(2);
            cell.setCellStyle(columnCellStyle);
            cell.setCellValue("Parent Layer ID");

            cell = layers.createCell(3);
            cell.setCellStyle(columnCellStyle);
            cell.setCellValue("Visible");

            cell = layers.createCell(4);
            cell.setCellStyle(columnCellStyle);
            cell.setCellValue("Locked");

            layers = sheet.createRow(rowIndex++);
            cell = layers.createCell(0);
            cell.setCellValue(200);

            cell = layers.createCell(1);
            cell.setCellValue("");

            cell = layers.createCell(2);
            cell.setCellValue("");

            cell = layers.createCell(3);
            cell.setCellValue(true);

            cell = layers.createCell(4);
            cell.setCellValue(false);

            layers = sheet.createRow(rowIndex++);
            cell = layers.createCell(0);
            cell.setCellValue(201);

            cell = layers.createCell(1);
            cell.setCellValue("Background");

            cell = layers.createCell(2);
            cell.setCellValue(200);

            cell = layers.createCell(3);
            cell.setCellValue(true);

            cell = layers.createCell(4);
            cell.setCellValue(false);


            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);

            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Map<String ,ArrayList<Node>> splitByType(ArrayList<Node> nodes){
        Map<String, ArrayList<Node>> mapByType = new HashMap<>();

        for(Node node : nodes){
            mapByType.computeIfAbsent(node.getType(), k -> new ArrayList<>()).add(node);
        }
        return mapByType;
    }

    public List<Node> handleFileUpload(MultipartFile file) throws IOException {
        if (file.isEmpty()) {
            throw new IOException("Empty file");
        }

        List<Node> nodeList = new ArrayList<>();
        String currentType = null;

        try {
            InputStream is = file.getInputStream();
            Workbook workbook = WorkbookFactory.create(is);
            Sheet sheet = workbook.getSheetAt(0);

            // Start processing from row 13 (index 12), end at without the last 3 rows since those are the layers, which are default data
            for (int i = 12; i <= sheet.getLastRowNum() - 3; i++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                Cell typeCell = row.getCell(0);
                if(typeCell == null) continue;
                if (typeCell.getCellType() == CellType.STRING) {
                    if (Objects.equals(typeCell.getStringCellValue(), "Sources")||Objects.equals(typeCell.getStringCellValue(), "Drains")||Objects.equals(typeCell.getStringCellValue(), "Pools")||Objects.equals(typeCell.getStringCellValue(), "Gates")||Objects.equals(typeCell.getStringCellValue(), "Traders")||Objects.equals(typeCell.getStringCellValue(), "Converters")||Objects.equals(typeCell.getStringCellValue(), "End Conditions")||Objects.equals(typeCell.getStringCellValue(), "Delays")||Objects.equals(typeCell.getStringCellValue(), "Texts")||Objects.equals(typeCell.getStringCellValue(), "Registers")||Objects.equals(typeCell.getStringCellValue(), "Resource Connections")||Objects.equals(typeCell.getStringCellValue(), "State Connections")) {
                        currentType = typeCell.getStringCellValue();
                        continue;
                    }
                }
                Node node = new Node();
                if(row.getCell(0).getColumnIndex() == 0 && row.getCell(0).getCellType() != CellType.NUMERIC) continue;

                for (Cell cell : row) {
                    switch (currentType){
                        case "Sources" -> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Source");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Drains"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Drain");

                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Pools"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Pool");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Gates"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Gate");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Traders"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Trader");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Converters"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Converter");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Registers"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Register");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "Delays"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("Delay");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 6 -> node.setActivationMode(cell.getStringCellValue());
                                default -> {
                                }
                            }
                        }
                        case "End Conditions"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("EndCondition");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 4 -> {
                                    int[] position = extractPositionFromGeometry(cell.getStringCellValue());
                                    node.setX(position[0]);
                                    node.setY(position[1]);
                                }
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                default -> {
                                }
                            }
                        }
                        case "Resource Connections"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("ResourceConnection");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 8 -> node.setStartId((long) cell.getNumericCellValue());
                                case 9 -> node.setEndId((long) cell.getNumericCellValue());
                                default -> {
                                }
                            }
                        }
                        case "State Connections"-> {
                            switch (cell.getColumnIndex()) {
                                case 0 -> {
                                    node.setNodeId((long) cell.getNumericCellValue());
                                    node.setType("StateConnection");
                                }
                                case 1 -> node.setLabel(cell.getStringCellValue());
                                case 5 -> {
                                    String color = cell.getStringCellValue().split(";")[8].split("=")[1];
                                    String size = cell.getStringCellValue().split(";")[3].split("=")[1];
                                    node.setColor(color);
                                    node.setSize(Integer.parseInt(size));
                                }
                                case 7 -> node.setStartId((long) cell.getNumericCellValue());
                                case 8 -> node.setEndId((long) cell.getNumericCellValue());
                                default -> {
                                }
                            }
                        }
                    }
                }
                if(node.getNodeId() != null) {
                    nodeList.add(node);
                }
            }

            return nodeList;
        } catch (Exception e) {
            e.printStackTrace();
            throw new IOException("Error writing to file: " + e);
        }
    }

    private int[] extractPositionFromGeometry(String geometry) {
        JSONObject jsonGeometry = new JSONObject(geometry);
        int x = jsonGeometry.getInt("x");
        int y = jsonGeometry.getInt("y");
        return new int[]{x, y};
    }
}
