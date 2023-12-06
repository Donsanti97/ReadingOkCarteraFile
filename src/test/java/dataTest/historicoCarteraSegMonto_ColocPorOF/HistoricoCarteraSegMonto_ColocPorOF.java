package dataTest.historicoCarteraSegMonto_ColocPorOF;

import org.apache.poi.util.IOUtils;
import org.testng.annotations.Test;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.*;
import java.util.List;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.MethotsAzureMasterFiles.*;

public class HistoricoCarteraSegMonto_ColocPorOF {
    //110 hojas

    private static String menu(List<String> opciones) {

        JFrame frame = new JFrame("Menú de Opciones");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JComboBox<String> comboBox = new JComboBox<>(opciones.toArray(new String[0]));
        comboBox.setSelectedIndex(0);

        JButton button = new JButton("Seleccionar");

        ActionListener actionListener = e -> frame.dispose();

        button.addActionListener(actionListener);

        JPanel panel = new JPanel();
        panel.add(comboBox);
        panel.add(button);

        frame.add(panel);
        frame.setSize(300, 100);
        frame.setVisible(true);

        while (frame.isVisible()) {
            // Esperar hasta que la ventana se cierre
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return comboBox.getSelectedItem().toString();
    }

    public static void configuracion(String masterFile) {


        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();

        JOptionPane.showMessageDialog(null, "A continuación se creará un archivo temporal " +
                "\n Se recomienda seleccionar la carpeta \"Documentos\" para esta función...");
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";

        try {
            waitSeconds(10);
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitSeconds(5);



            nuevosOficinas(okCartera, tempFile);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);
        } catch (HeadlessException | IOException e) {
            throw new RuntimeException(e);
        }
    }


    @Test
    public static void test(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";
        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCartera);
        List<Map<String, String>> datosFiltrados;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        for (String sheetName :
                sheetNames) {
            System.out.println("SheetName: " + sheetName);
            sheetName = "Hoja1";

            List<String> encabezados = obtenerEncabezados(okCartera, sheetName);
            for (String encabezado : encabezados) {
                System.out.println("Header: " + encabezado);
            }

            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCartera, sheetName, campoFiltrar, valorInicio, valorFin, tempFile);
            System.out.println("DATOS_FILTRADOS: " + datosFiltrados.size() + " : " + datosFiltrados);


            sheetNames = obtenerNombresDeHojas(tempFile);

            for (Map<String, String> rowData : datosFiltrados) {
                System.out.println(rowData);
                for (String campoDeseado : camposDeseados) {
                    if (rowData.containsKey(campoDeseado)) {
                        String valorCampo = rowData.get(campoDeseado);
                        if (valorCampo != null) {
                            System.out.println(campoDeseado + ": " + valorCampo);
                        } else {
                            System.out.println(campoDeseado + ": Valor nulo o campo no encontrado");
                        }
                    } else {
                        System.out.println(campoDeseado + ": Campo no encontrado");
                    }
                }
                System.out.println();
            }
        }

    }

    public static void nuevosOficinas(String okCarteraFile/*, String masterFile, String azureFile, String fechaCorte, String hoja*/, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        //try {


            List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

            List<String> headers;
            List<Map<String, String>> datosFiltrados = null;
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            for (String sheetName : sheetNames) {
                System.out.println("Contenido de la hoja: " + sheetName);
                headers = obtenerEncabezados(okCarteraFile, sheetName);

                // Listar campos disponibles
                System.out.println("Campos disponibles:");
                for (String header : headers) {
                    System.out.println(header);
                }
                // Especifica el campo en el que deseas aplicar el filtro
                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, tempFile);


                // Especifica los campos que deseas obtener


                // Imprimir datos filtrados
                System.out.println("DATOS FILTRADOS");
                for (Map<String, String> rowData : datosFiltrados) {
                    for (String campoDeseado : camposDeseados) {
                        String valorCampo = rowData.get(campoDeseado);

                        System.out.println(campoDeseado + ": " + valorCampo);
                    }
                    System.out.println();
                }
                runtime();
                waitSeconds(2);


            }
            System.out.println("-----------CREACIÓN TEMPORAL-----------");

            // Crear una nueva hoja Excel con los datos filtrados
            crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

            /*System.out.println("Análisis archivo temporal----------------------");

            sheetNames = obtenerNombresDeHojas(tempFile);

            for (String sheetName : sheetNames) {
                System.out.println("Contenido de la hoja: " + sheetName);

                headers = obtenerEncabezados(tempFile, sheetName);

                System.out.println("Campos disponibles " + headers);

                for (String header : headers) {
                    System.out.println(header);
                }


                //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
                datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


                System.out.println("VALORES DEL OK CARTERA");
                for (Map<String, String> rowData : datosFiltrados) {
                    for (String campoDeseado : camposDeseados) {
                        String valorCampo = rowData.get(campoDeseado);
                        System.out.println(campoDeseado + ": " + valorCampo);
                    }
                    System.out.println();
                }
                runtime();
                waitSeconds(2);

                Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

                List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)*//*getHeadersMFile(azureFile, masterFile, fechaCorte)*//*;


                for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                    for (Map<String, String> datoMF : datosMasterFile) {
                        for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                            *//*------------------------------------------------------------*//*
                            if (entryOkCartera.getKey().contains(entry.getKey())) {

                                System.out.println("CÓDIGO ENCONTRADO");


                                if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                    System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CÓDIGO: " + entry.getKey());
                                } else {

                                    System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CÓDIGO: " + entry.getKey());

                                }
                            } else {
                                System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                            }
                            *//*-------------------------------------------------------------------*//*
                        }
                        waitSeconds(5);
                        runtime();
                    }

                }
                runtime();

            }*/
        /*} catch (IOException e) {
            throw new RuntimeException("Error interno del proceso", e);
        }*/
    }


}
