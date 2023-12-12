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
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.List;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;



import static org.dataTest.FunctionsApachePoi.*;
import static org.dataTest.MethotsAzureMasterFiles.*;

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
            List<Map<String, Object>> values = getHeaderFilterValuesN(sheet, headers, "tipo_cliente", "Nuevo", "Nuevo");
            for (Map<String, Object> rowData : values) {
                for (String fields : requiredFields) {
                    if (rowData.containsKey(requiredFields)) {
                        String value = (String) rowData.get(fields);
                        if (value != null) {
                            System.out.println(fields + ": " + value);
                        } else {
                            System.err.println("Valor nulo o campo no encontrado");
                        }
                    }else {
                        System.out.println(fields + ": Campo no encontrado");
                    }

                }
                System.out.println();
            }

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private static  List<String> getHeadersN(Sheet sheet){
        List<String> columnNames = new ArrayList<>();
        Row row = sheet.getRow(0);
        try {
            System.out.println("PROCESANDO CAMPOS...");
            for (Iterator<Cell> it = row.cellIterator(); it.hasNext(); ) {
                Cell cell = it.next();
                columnNames.add(obtenerValorVisibleCelda(cell));
                //System.out.println(obtenerValorVisibleCelda(cell));
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return columnNames;
    }

    private static Map<String, Object> getHeaderValuesN(Sheet sheet, List<String> headers){
        Map<String, Object> rowData = new HashMap<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        try {
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }
                Iterator<String> columnNameIterator = headers.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String columnName = columnNameIterator.next();
                    String value = "";
                    if (cell != null) {
                        //System.out.println("The cell contains a numeric value." + cell.getCellType());
                        value = obtenerValorVisibleCelda(cell);
                        System.out.println("VALUE: " + value + ", " + cell.getRowIndex());
                        rowData.put(columnName, value);
                    }
                    runtime();
                    Thread.sleep(200);
                    //waitSeconds(2);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }


        return rowData;
    }

    private static List<Map<String, Object>> getHeaderFilterValuesN(Sheet sheet, List<String> headers, String campoFiltrar, String valorIni, String valorFin){
        List<Map<String, Object>> datosFiltrados = new ArrayList<>();

        Row row = sheet.getRow(0);

        Iterator<Row> rowIterator = sheet.iterator();

        int totalRows = sheet.getPhysicalNumberOfRows() - 1;

        try {
            int currentRow = 0;
            System.out.println("PROCESANDO VALORES");
            while (rowIterator.hasNext()) {

                row = rowIterator.next();
                if (row.getRowNum() == 0) {
                    continue;
                }
                int campoFiltrarIndex = headers.indexOf(campoFiltrar);
                if (campoFiltrarIndex == -1) {
                    System.err.println("El campo especificado para el filtro no existe");
                    return  datosFiltrados;
                }

                String valueCampoFiltrar = obtenerValorVisibleCelda(row.getCell(campoFiltrarIndex));
                Iterator<String> columnNameIterator = headers.iterator();
                Iterator<Cell> cellIterator = row.cellIterator();
                if (valueCampoFiltrar.compareTo(valorIni) >= 0 && valueCampoFiltrar.compareTo(valorFin) <= 0) {

                    Map<String, Object> rowData = new HashMap<>();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String columnName = columnNameIterator.next();
                        String value = "";
                        if (cell != null) {
                            //System.out.println("The cell contains a numeric value." + cell.getCellType());
                            value = obtenerValorVisibleCelda(cell);
                            //System.out.println("VALUE: " + value + ", " + cell.getRowIndex());
                            rowData.put(columnName, value);
                        }
                        runtime();
                        Thread.sleep(200);
                        //waitSeconds(2);
                    }
                    datosFiltrados.add(rowData);
                    currentRow++;

                    showProgressBar(currentRow, totalRows);

                    Thread.sleep(50);
                }
            }
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }

        return datosFiltrados;
    }

    private static void showProgressBar(int current, int total) {
        int progressBarWidth = 50;
        int progress = (int) ((double) current / total * 100);

        StringBuilder progressBar = new StringBuilder("[");
        for (int i = 0; i < progressBarWidth; i++) {
            if (i < progress * progressBarWidth / 100) {
                progressBar.append("=");
            } else {
                progressBar.append(" ");
            }
        }
        progressBar.append("] " + progress + "%");
        System.out.print("\r" + progressBar.toString());
    }



    public static void test(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        waitSeconds(3);
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";
        IOUtils.setByteArrayMaxOverride(300000000);

        //List<String> sheetNames = obtenerNombresDeHojas(okCartera);



        List<Map<String, String>> datosFiltrados;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");


        try {
            Workbook workbook = WorkbookFactory.create(new File(okCartera));
            Sheet sheet = workbook.getSheetAt(0);
            //for (String sheetName : sheetNames) {
                System.out.println("SheetName: " + sheet.getSheetName());
                //sheetName = "Hoja1";

                List<String> encabezados = getHeadersN(sheet/*, sheet.getSheetName()*/);
                for (String encabezado : encabezados) {
                    System.out.println("Header: " + encabezado);
                }

                String campoFiltrar = "tipo_cliente";
                String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
                String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango

                // Filtrar los datos por el campo y el rango especificados
                datosFiltrados = obtenerValoresDeEncabezados(okCartera, sheet.getSheetName(), campoFiltrar, valorInicio, valorFin);
                System.out.println("DATOS_FILTRADOS: " + datosFiltrados.size() + " : " + datosFiltrados);


                //sheetNames = obtenerNombresDeHojas(tempFile);
            OPCPackage opcPackage = OPCPackage.open((new File( tempFile)));
            workbook =  new XSSFWorkbook(new File(tempFile));
            Sheet sheet1 = workbook.getSheetAt(0);

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
            //}
        } catch (Exception e) {
            throw new RuntimeException(e);
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
