package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Set;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");

        for (Client client : allClients) {
            int Age = now.getYear() - client.getDateNaissance().getYear();
            writer.println(client.getId() + ";"
                    + "\"" + client.getNom() + "\"" + ";"
                    + "\"" + client.getPrenom() + "\"" + ";"
                    + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")) + ";"
                    + Age);
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsxlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setContentType("application/vnd.ms-excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellNom = headerRow.createCell(1);
        cellNom.setCellValue("Nom");

        Cell cellPrenom = headerRow.createCell(2);
        cellPrenom.setCellValue("Prénom");

        Cell cellDateNaissance = headerRow.createCell(3);
        cellDateNaissance.setCellValue("Date de naissance");

        Cell cellAge = headerRow.createCell(4);
        cellAge.setCellValue("Age");

        int i = 0;
        for (Client client : allClients) {

            int Age = now.getYear() - client.getDateNaissance().getYear();
            Row row = sheet.createRow(i + 1);

            Cell cellIdClient = row.createCell(0);
            cellIdClient.setCellValue(client.getId());

            Cell cellNomClient = row.createCell(1);
            cellNomClient.setCellValue(client.getNom());

            Cell cellPrenomClient = row.createCell(2);
            cellPrenomClient.setCellValue(client.getPrenom());

            Cell cellDateNaissanceClient = row.createCell(3);
            cellDateNaissanceClient.setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));

            Cell cellAgeClient = row.createCell(4);
            cellAgeClient.setCellValue(Age);

            i++;

        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void facturesXLXSByclient(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException {

        response.setContentType("application/vnd.ms-excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client" + clientId + ".xlsx\"");
        List<Facture> factures = factureService.findFacturesClient(clientId);


        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Factures");

        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellPrixTotal = headerRow.createCell(1);
        cellPrixTotal.setCellValue("Prix total");

        int i = 0;
        for (Facture facture : factures) {

            Row row = sheet.createRow(i + 1);

            Cell cellIdFacture = row.createCell(0);
            cellIdFacture.setCellValue(facture.getId());

            Cell cellPrixTotalFacture = row.createCell(1);
            cellPrixTotalFacture.setCellValue(facture.getTotal());

            i++;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesxlsx(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel\n");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        List<Client> clients = clientService.findAllClients();

        Workbook workbook = new XSSFWorkbook();
        for (Client client : clients) {
            Sheet sheet = workbook.createSheet(client.getNom());

            Row headerRow = sheet.createRow(0);

            Cell cellNomClient = headerRow.createCell(0);
            cellNomClient.setCellValue(client.getNom());

            Cell cellPrenomClient = headerRow.createCell(1);
            cellPrenomClient.setCellValue(client.getPrenom());

            List<Facture> factures = factureService.findFacturesClient(client.getId());
            for (Facture facture : factures) {
                Sheet sheetFacture = workbook.createSheet("Facture " + facture.getId());

                int i=0;
                for(LigneFacture ligneFacture : facture.getLigneFactures()){
                    Row headerRowFacture = sheetFacture.createRow(i);

                    Cell cellNomArticleHeader = headerRowFacture.createCell(0);
                    cellNomArticleHeader.setCellValue("Libellé article");
                    Cell cellQuantiteHeader = headerRowFacture.createCell(1);
                    cellQuantiteHeader.setCellValue("Quantité commandée");
                    Row rowFacture = sheetFacture.createRow(1);

                    Cell cellNomArticle = rowFacture.createCell(0);
                    cellNomArticle.setCellValue(ligneFacture.getArticle().getLibelle());
                    Cell cellQuantite = rowFacture.createCell(1);
                    cellQuantite.setCellValue(ligneFacture.getQuantite());
                    Cell cellPrixUnitaire = rowFacture.createCell(2);
                    cellPrixUnitaire.setCellValue(ligneFacture.getArticle().getPrix());
                    Cell cellPrixLigne = rowFacture.createCell(3);
                    cellPrixLigne.setCellValue(ligneFacture.getSousTotal());

                    i++;
                }
            }
        }
        workbook.write(response.getOutputStream());
        workbook.close();
        }
    }
