package com.example.MachinationsServer.Conroller;

import com.example.MachinationsServer.Models.Diagram;
import com.example.MachinationsServer.Models.Node;
import com.example.MachinationsServer.Service.DiagramService;

import org.springframework.beans.factory.annotation.Autowired;

import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@CrossOrigin(origins = "http://localhost:*")
@RestController
@RequestMapping("/machinations_api")
public class MachinationsController {

    @Autowired
    DiagramService service;

    @PostMapping("/save/{name}")
    public void saveSvg(@RequestBody ArrayList<Node> nodes, @PathVariable String name) throws IOException {
        System.out.println(nodes);
        service.saveDiagram(name, nodes);
    }

    @PostMapping("/save_csv/{name}")
    public void saveCSV(@RequestBody ArrayList<Node> diagramProperites, @PathVariable String name) throws IOException {
        service.saveToXLSX(name,diagramProperites);
    }

    @GetMapping("/load_all")
    public List<Diagram> loadDiagrams(){
        return service.getAllDiagrams();
    }

    @PostMapping("/send_file")
    public List<Node> getDiagramFromFile(@RequestBody MultipartFile file) throws IOException {
        List<Node> list = service.handleFileUpload(file);
        System.out.println(list);
        return list;
    }

    @GetMapping(value = "/")
    public String testConnection() {
        return "Connection established!";
    }

}