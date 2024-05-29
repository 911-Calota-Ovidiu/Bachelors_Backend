package com.example.MachinationsServer.Repository;

import com.example.MachinationsServer.Models.Diagram;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.Optional;

public interface IDiagramRepo extends JpaRepository<Diagram,Long> {
    Optional<Diagram> findByName(String name);
}
