/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */

package com.automacao_excel.automacao;

import java.io.IOException;

/**
 *
 * @author anvnc
 */
public class Main {
    public static void main(String[] args) throws IOException {
        DirectoryPaths searchPath = new DirectoryPaths();
        AbrirArquivo analisarISS = new AbrirArquivo();
        analisarISS.verPlanilha(searchPath);
        searchPath.getMainFrame().dispose();
        System.exit(0);
    }
}
