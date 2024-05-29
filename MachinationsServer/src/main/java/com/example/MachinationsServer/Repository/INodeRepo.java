package com.example.MachinationsServer.Repository;

import com.example.MachinationsServer.Models.Node;
import org.springframework.data.jpa.repository.JpaRepository;

public interface INodeRepo extends JpaRepository<Node,Long> {
}
