package org.dataTest;



import org.dataTest.historicoCarteraSegMonto_ColocPorOF.HistoricoCarteraSegMonto_ColocPorOF;

import javax.swing.*;
import java.io.File;

import static org.dataTest.MethotsAzureMasterFiles.getDocument;

public class Start {

    public static void excecution(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        String masterFile = getDocument();

        try {
            assert masterFile != null;
            File file = new File(masterFile);
            System.out.println(file.getName());
            String fileName = file.getName().toLowerCase();
            System.out.println(fileName);
            if (fileName.contains("historico cartera seg monto_coloc por of")){
                HistoricoCarteraSegMonto_ColocPorOF.configuracion(masterFile);
            }else {
                System.out.println("EL ARCHIVO SELECCIONADO NO TIENE ANÁLISIS ASIGNADO");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
}
