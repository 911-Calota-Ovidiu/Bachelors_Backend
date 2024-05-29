package com.example.MachinationsServer.Models;

import com.fasterxml.jackson.annotation.JsonIgnore;
import jakarta.persistence.*;
import lombok.*;

@SuppressWarnings("JpaDataSourceORMInspection")
@Entity
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
@Table(name = "node")
public class Node {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "diagram_id", nullable = false)
    @JsonIgnore
    @ToString.Exclude
    @EqualsAndHashCode.Exclude
    private Diagram diagram;

    @Column(name = "nodeId")
    private Long nodeId;

    @Column(name = "type")
    private String type;

    @Column(name = "x")
    private Integer x;

    @Column(name = "y")
    private Integer y;

    @Column(name = "color")
    private String color;

    @Column(name = "size")
    private  Integer size;

    @Column(name = "label")
    private String label;

    @Column(name = "activationMode")
    private String activationMode;

    @Column(name = "startId")
    private Long startId;

    @Column(name = "endId")
    private Long endId;

}
