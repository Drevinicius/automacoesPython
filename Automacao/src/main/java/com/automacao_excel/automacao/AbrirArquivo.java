/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.automacao_excel.automacao;

import javax.swing.JOptionPane;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import java.util.List;
import java.util.ArrayList;
import java.text.DecimalFormat;

/**
 *
 * @author anvnc
 * Essa classe abre e manipula meu arquivo
 */
public class AbrirArquivo {
    
    String fileName;
    private final DataFormatter str = new DataFormatter(); // Definir meu cell type sempre para String
    DecimalFormat df = new DecimalFormat("00000000");
    
    // Convesor de datas de MM/dd/yyyy para dd/MM/yyyy
    private String converterData(String dataEntrada){
        String[] dataDividida = dataEntrada.split("/");
        String dataFormatada = "";
        String auxiliar = "";
        if(dataDividida[1].length() < 2){
            auxiliar = "0" + dataDividida[1];
            dataFormatada = auxiliar;
        }else{
            dataFormatada = dataDividida[1];
        }if(dataDividida[0].length() < 2){
            auxiliar = "0" + dataDividida[0];
            dataFormatada = dataFormatada + "/" + auxiliar;
        }else{
            dataFormatada = dataFormatada + "/" + dataDividida[0];
        }if(dataDividida[2].length()< 3){
            auxiliar = "20" + dataDividida[2];
            dataFormatada = dataFormatada + "/" + auxiliar;            
        }else{
            dataFormatada = dataFormatada + "/" + dataDividida[2];
        }
        
        return dataFormatada;
    }

    // Metodo conversor para limpar 0s a mais nos meus numeros flutuantes
    private static String conversor(String valor){
        String[] listaString = valor.split(",");
        String nova = "";
        if(Integer.parseInt(listaString[1]) == 0){
            nova = listaString[0];
        }else if(Integer.parseInt(listaString[1]) < 10){
            nova = valor.replace(",", ".");
        }else{
            String replace = listaString[1].replace("0", "");
            nova = listaString[0] + "." + replace;
        }
        
        return nova;
    }
    
    // Metodo de extração de dados
    private List<List<String>> extrairDados(Sheet sheet){
        List<List<String>> dados = new ArrayList<>(); // Crio o array de arrays do tipo String que guardara meus dados
        
        for(Row row: sheet){ // Pega todas as linhas da minha pagina
            List<String> linha = new ArrayList<>();
            for(Cell cell: row){ // Pega todas as celulas da minha linha
                linha.add(str.formatCellValue(cell)); // Adiciono minha cell value como String
            }
            dados.add(linha); // Adiciono ao meu array de arrays
        }

        return dados;
    }
    // metodo 'MAIN' que abre e faz todas as analises necessárias para salvar meu arquivo
    public void verPlanilha(DirectoryPaths pathMain){
        String path = pathMain.setPathFile("Selecione a planilha do ISS");
        try(FileInputStream fileInISS = new FileInputStream(path)){
            fileName = pathMain.getPathFile().getName();
            try(FileInputStream fileInNotas = new FileInputStream(pathMain.setPathFile("Selecione a planilha das notas"))){
                // Acessando as 2 planilha solicitadas do usuario
                Workbook planilhaDoISS = new XSSFWorkbook(fileInISS);
                Workbook planilhaDasNotas = new XSSFWorkbook(fileInNotas);
                
                // Minhas páginas selecionadas das minhas planilhas
                Sheet paginaDoISS;
                Sheet novaPaginaDoISS;
                Sheet paginaDasNotas;
                
                // Listas compostas que vão receber outras listas do tipo List<String> dos dados extraídos das minhas planilhas
                List<List<String>> notasLancadas = new ArrayList<>();
                List<List<String>> notasISS = new ArrayList<>();
                List<List<String>> notasNaoEncontradas = new ArrayList<>();
                
                // Verificando se na pllanilha do ISS existe uma página 'Notas analisadas' se não cria
                if(planilhaDoISS.getSheetIndex("Notas analisadas") < 0){
                    novaPaginaDoISS = planilhaDoISS.createSheet("Notas analisadas");
                    paginaDoISS = planilhaDoISS.getSheetAt(0); // Usado para fazer comparação com as notas
                }else{
                    paginaDoISS = planilhaDoISS.getSheetAt(0); // Usado para fazer comparação com as notas
                    planilhaDoISS.removeSheetAt(planilhaDoISS.getSheetIndex("Notas analisadas"));
                    novaPaginaDoISS = planilhaDoISS.createSheet("Notas analisadas"); // Recebe a comparação
                }
                
                paginaDasNotas = planilhaDasNotas.getSheetAt(0); // Usado para comparação com o ISS
                
                // Chamada do metodo que extrai os dados das planilhas
                notasLancadas = extrairDados(paginaDasNotas);
                notasISS = extrairDados(paginaDoISS);
                
                // Varíaveis que vão servi de comparação dos itens da minha List<List<String>>
                String nfISS = "0";
                String emissaoISS = "a";
                String valorDoISS = "a";
  
                
                String nfSIS = "1";
                String emissaoSIS = "b";
                String valorDoSIS = "b";
          
                boolean encontrou = false; // Caso a nota exista na planilha ele será verificado
                
                for(List<String> iss: notasISS){
                    encontrou = false; // Sempre considero que não exista
                    if(iss.size() >= 11){
                        emissaoISS = iss.get(3).trim();
                        
                        valorDoISS = iss.get(7).trim().replace(".", "");
                        valorDoISS = valorDoISS.replace(",", ".");
                 
                        try{
                            nfISS = iss.get(1).trim();
                            
                            nfISS = df.format(Integer.parseInt(nfISS));
                            nfISS = nfISS.replaceAll("[^0-9]", "");
                        }catch(NumberFormatException e){
                            nfISS = "00000000";
                        }
                        for(List<String> sis: notasLancadas){
                            if(sis.size() > 41){
                                emissaoSIS = sis.get(5).trim();
                                valorDoSIS = sis.get(16).trim().replace(".", "");

                                try{
                                    nfSIS = sis.get(3).trim();
                                    nfSIS = df.format(Integer.parseInt(nfSIS));
                                    nfSIS = nfSIS.replaceAll("[^0-9]", "");
                                }catch(NumberFormatException e){
                                    nfSIS = "000000000";
                                }


                                if(nfISS.equalsIgnoreCase(nfSIS)){
                                    if(emissaoISS.equals(converterData(emissaoSIS))){
                                        if(valorDoISS.equals(conversor(valorDoSIS))){
                                            encontrou = true; // Caso ache troco a variavel
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        // Se não achar inverto a varíavel a adiciono aos itens não encontrados
                        if(!encontrou){
                            notasNaoEncontradas.add(iss);
                        }
                    }
                }
                // Faço um laço para adicionar a minha nova planilha
                int interador = 0;
                for(List<String> lista: notasNaoEncontradas){
                    Row novaLinha = novaPaginaDoISS.createRow(interador++);
                    for(int i = 0; i < lista.size(); i++){
                        Cell cell = novaLinha.createCell(i);
                        cell.setCellValue((String) lista.get(i));
                    }
                }
                
                try(FileOutputStream fileOut = new FileOutputStream(path)){
                    planilhaDoISS.write(fileOut);
                }catch(IOException e){
                    e.printStackTrace();
                }finally{
                    planilhaDoISS.close();
                    planilhaDasNotas.close();
                }
            }catch(IOException e){
                    JOptionPane.showMessageDialog(null,
                    "Selecione um arquivo excel válido",
                    "Arquivo não encontrado",
                    JOptionPane.ERROR_MESSAGE);
            }
        }catch(IOException e){
            JOptionPane.showMessageDialog(null,
                    "Selecione um arquivo excel válido",
                    "Arquivo não encontrado",
                    JOptionPane.ERROR_MESSAGE);
        }finally{
            JOptionPane.showMessageDialog(pathMain.getMainFrame(), 
                    "Arquivo " + fileName + " salvo" , 
                    "Processo finalizado", 
                    JOptionPane.INFORMATION_MESSAGE);
            pathMain.getMainFrame().dispose();
            System.exit(0);
        }
    }
}
