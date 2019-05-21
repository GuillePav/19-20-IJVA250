package com.example.demo.repository;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Set;

@Repository
public interface FactureRepository extends JpaRepository<Facture, Long> {

    //@Query("SELECT f FROM Facture WHERE f.client.id = ?1"
    List<Facture> findByClientId(Long clientId);

}
