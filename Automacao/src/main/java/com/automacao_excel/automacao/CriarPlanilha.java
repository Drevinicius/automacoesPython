/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.automacao_excel.automacao;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;

/**
 *
 * @author anvnc
 */
public final class CriarPlanilha {
    public void novaPlanilha(DirectoryPaths pathName){
        String pathNameDirectory = pathName.setPathDirectory();
        Workbook planilha = new XSSFWorkbook();
        Sheet cliente = planilha.createSheet("Cliente");

        Row header = cliente.createRow(0);
        String[] cabecalho = {"ID", "NOME", "EMAIL", "CADASTRO"};
        Object[][] objetos = {
            {"1", "ANDRE", "ANVNC01@GMAIL.COM", "01/02/2026"},
            {"2", "ANDRE", "ANVNC01@GMAIL.COM", "01/02/2026"},
            {"3", "ANDRE", "ANVNC01@GMAIL.COM", "01/02/2026"}};

        for(int i = 0; i < cabecalho.length; i++){
            Cell novaLinha = header.createCell(i);
            novaLinha.setCellValue(cabecalho[i]);
        }

        int interador = 1;
        for(Object[] linha: objetos){
            Row novaLinha = cliente.createRow(interador++);

            for (int i = 0; i < linha.length; i++) {
                Cell cell = novaLinha.createCell(i);
                if (linha[i] instanceof String) {
                    cell.setCellValue((String) linha[i]);
                } else if (linha[i] instanceof Integer) {
                    cell.setCellValue((Integer) linha[i]);
                }
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(pathNameDirectory)) {
            planilha.write(fileOut);
            //System.out.println("Arquivo Excel criado com sucesso!");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                planilha.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }    
}
