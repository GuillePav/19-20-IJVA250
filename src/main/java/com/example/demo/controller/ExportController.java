package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.service.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id;Nom;Prenom;Date de Naissance;Age");

        for(Client client:allClients){
            int Age =  now.getYear() - client.getDateNaissance().getYear();
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

        int i=0;
        for (Client client : allClients) {

            int Age =  now.getYear() - client.getDateNaissance().getYear();
            Row row = sheet.createRow(i+1);
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
}
