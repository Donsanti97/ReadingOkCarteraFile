package org.dataTest.historicoCarteraSegMonto_ColocPorOF;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.List;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;



import static org.dataTest.FunctionsApachePoi.*;
import static org.dataTest.MethotsAzureMasterFiles.*;

public class HistoricoCarteraSegMonto_ColocPorOF {
    //110 hojas

    public static String menu(List<String> opciones) {

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

    public static void testWithJson(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        waitSeconds(3);
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";
        IOUtils.setByteArrayMaxOverride(300000000);

        try {
            //String excelFilePath = "C:\\Users\\01925\\Downloads\\prueba.xlsx";

            Workbook workbook = WorkbookFactory.create(new File(okCartera));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            JsonArray carArray = new JsonArray();
            List<String> columnNames = new ArrayList<>();

            Gson gson = new Gson();

            // Get column names
            System.err.println("COLUMN NAMES");
            Row row = sheet.getRow(0);
            for (Iterator<Cell> it = row.cellIterator(); it.hasNext(); ) {
                Cell cell = it.next();
                columnNames.add(obtenerValorVisibleCelda(cell));
                System.out.println(obtenerValorVisibleCelda(cell));
            }


            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum()==0) {
                    continue;
                }
                Iterator<String> columnNameIterator = columnNames.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();

                // Create a new map for the row
                Map<String, Object> newCarMap = new HashMap<>();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String columnName = columnNameIterator.next();
                    String value = "";
                    if (cell/*.getCellType()*/!=null) {
                        System.out.println("The cell contains a numeric value."+cell.getCellType());
                        value = obtenerValorVisibleCelda(cell);
                        System.out.println("VALUE: " + value +", "+ cell.getRowIndex());
                        newCarMap.put(columnName, value);
                    }
                    runtime();
                    Thread.sleep(200);
                    //waitSeconds(2);
                }
                // Convert the map to `JsonElement`
                JsonElement carJson = gson.toJsonTree(newCarMap);
                // Add the `JsonElement` to `JsonArray`
                carArray.add(carJson);
            }
            // Add the `JsonArray` to `completeJson` object with the key as `Cars`
            JsonObject completeJson = new JsonObject();



            completeJson.add("codigo_sucursal", carArray);

        } catch (IOException ex) {
            throw new RuntimeException(ex);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    public static void testWithNew(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        waitSeconds(3);
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";
        IOUtils.setByteArrayMaxOverride(300000000);

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCartera));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> requiredFields = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<Map<String, Object>> values = getHeaderFilterValuesNS(sheet, headers, "tipo_cliente", "Nuevo", "Nuevo");

            int currentRow = 0;
            int rowsPerBatch = 5000;

            for (Map<String, Object> rowData : values) {
                for (String fields : requiredFields) {

                    for (Map.Entry<String, Object> entry : rowData.entrySet()){
                        if (entry.getKey().contains(fields)) {
                            System.out.println("CAMPO: " + entry.getKey() + " - VALOR: " + entry.getValue());
                        }/*else {
                            System.out.println("ESTA VUELTA NO COINCIDE EN " + fields);
                        }*/
                    }

                    /*if (rowData.containsKey(requiredFields)) {
                        String value = (String) rowData.get(fields);
                        if (value != null) {
                            System.out.println(fields + ": " + value);
                        } else {
                            System.err.println("Valor nulo o campo no encontrado");
                        }
                    }else {
                        System.out.println(fields + ": Campo no encontrado");
                    }*/

                    currentRow++;

                    if (currentRow % rowsPerBatch == 0) {
                        runtime();
                        Thread.sleep(200);
                    }

                }
                System.out.println();
            }

            crearNuevaHojaExcel(requiredFields, values, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            //headers = getHeaders(sheet);

            //values = getHeaderValuesN(sheet, headers);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, requiredFields.get(0), requiredFields.get(1));

            int count = 0;
            System.out.println(" RESULTADO SUMATORIA OKAY_CARTERA");
            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                System.out.println("CAMPO: " + entryOkCartera.getKey() + ", VALORES: " + entryOkCartera.getValue());
                count++;
            }

            if (count % rowsPerBatch == 0) {
                runtime();
                Thread.sleep(200);
            }


        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

    }

    public static void testWithNewMasterFile(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Master");
        String masterFile = getDocument();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = mostrarCuadroDeTexto();

        try {
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, "Nuevos_Oficinas", fechaCorte);

            int count = 0;
            int rowsPerBatch = 5000;

            for (Map<String, String> data : datosMasterFile){
                for (Map.Entry<String, String> entry : data.entrySet()){
                    System.out.println("KEY: " + entry.getKey() + ", VALUE: " + entry.getKey());
                }
                count++;
            }

            if (count % rowsPerBatch == 0) {
                runtime();
                Thread.sleep(200);
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }


    }

    /*---------------------------------------------------------------------------------------------------------------------------------------*/

/*---------------------------------------------------------------------------------------------------------------------------------------*/
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
                datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);


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
