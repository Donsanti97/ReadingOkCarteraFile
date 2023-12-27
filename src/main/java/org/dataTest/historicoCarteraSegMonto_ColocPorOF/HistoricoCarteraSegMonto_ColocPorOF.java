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
import java.text.ParseException;
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

        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        /*JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        masterFile = getDocument();*/
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola el número del mes y año de corte del archivo OkCartera sin espacios (Ejemplo: 02/2023 (febrero/2023))");
        String mesAnoCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "A continuación se creará un archivo temporal " +
                "\n Se recomienda seleccionar la carpeta \"Documentos\" para esta función...");
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";

        try {
            waitSeconds(10);
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitSeconds(5);

            JOptionPane.showMessageDialog(null, "Para los análisis de algunas de las hojas a continuación es necesario que" +
                    "\n Digite a continuación un tipo de calificación entre [B] y [E]");
            List<String> opciones = Arrays.asList("B", "C", "D", "E");
            String calificacion = menu(opciones);

            nuevosOficinas(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas", tempFile);


            /*nuevosOficinasMay30(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas > 30", tempFile);

            nuevosOficinasBE(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas_B_E", calificacion, tempFile);

            renovadoOficinas(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas", tempFile);

            renovadoOficinasMay30(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas_>30", tempFile);

            renovadoOficinasBE(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas_B_E", calificacion, tempFile);

            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc '0-0.5 M", 0, 5, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 0.5-1 M", 5, 10, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 1-2 M", 10, 20, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 2-3 M", 20, 30, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 3-4 M", 30, 40, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 4-5 M", 40, 50, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 5-10 M", 50, 100, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 10-15 M", 100, 150, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 15-20 M", 150, 200, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 20-25 M", 200, 250, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 25-50 M", 250, 500, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 50-100 M", 500, 1000, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc > 100 M", 1000, 10000, tempFile);

            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol '0-0.5 >30", 0, 5, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 0.5-1 > 30", 5, 10, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 1-2M >30", 10, 20, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 2-3M >30", 20, 30, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 3-4M >30", 30, 40, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 4-5M >30", 40, 50, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 5-10M >30", 50, 100, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 10-15 >30", 100, 150, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 15-20 >30", 150, 200, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 20-25 >30", 200, 250, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 25-50 >30", 250, 500, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 50-100 >30", 500, 1000, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol > 100 >30", 1000, 10000, tempFile);

            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol '0-0.5 B_E", 0, 5, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 0.5-1 B_E", 5, 10, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 1-2 B_E", 10, 20, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 2-3 B_E", 20, 30, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 3-4 B_E", 30, 40, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 4-5 B_E", 40, 50, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 5-10 B_E", 50, 100, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 10-15 B_E", 100, 150, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 15-20 B_E", 150, 200, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 20-25 B_E", 200, 250, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 25-50 B_E", 250, 500, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 50-100 B_E", 500, 1000, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol > 100 B_E", 1000, 10000, calificacion, tempFile);

            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_'0-0.5 M", 0, 5, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_0.5-1 M", 5, 10, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_1-2M M", 10, 20, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_2-3M M", 20, 30, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_3-4M M", 30, 40, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_4-5M M", 40, 50, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_5-10M M", 50, 100, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_10-15 M", 100, 150, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_15-20 M", 150, 200, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_20-25 M", 200, 250, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_25-50 M", 250, 500, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_50-100 M", 500, 1000, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_> 100 M", 1000, 10000, tempFile);

            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_'0-0.5 M", 0, 5, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_0.5-1 M", 5, 10, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_1-2M M", 10, 20, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_2-3M M", 20, 30, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_3-4M M", 30, 40, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_4-5M M", 40, 50, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_5-10M M", 50, 100, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_10-15 M", 100, 150, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_15-20 M", 150, 200, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_20-25 M", 200, 250, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_25-50 M", 250, 500, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_50-100 M", 500, 1000, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_> 100 M", 1000, 10000, tempFile);

            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_'0-0.5 M", 0, 5, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_0.5-1 M", 5, 10, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_1-2M M", 10, 20, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_2-3M M", 20, 30, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_3-4M M", 30, 40, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_4-5M M", 40, 50, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_5-10M M", 50, 100, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_10-15 M", 100, 150, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_15-20 M", 150, 200, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_20-25 M", 200, 250, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_25-50 M", 250, 500, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_50-100 M", 500, 1000, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_> 100 M", 1000, 10000, tempFile);

            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_'0-0.5 M", 0, 5, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_0.5-1 M", 5, 10, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_1-2M M", 10, 20, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_2-3M M", 20, 30, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_3-4M M", 30, 40, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_4-5M M", 40, 50, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_5-10M M", 50, 100, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_10-15 M", 100, 150, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_15-20 M", 150, 200, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_20-25 M", 200, 250, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_25-50 M", 250, 500, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_50-100 M", 500, 1000, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_> 100 M", 1000, 10000, mesAnoCorte, tempFile);*/



            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);
        } catch (HeadlessException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    /*public static void testWithJson(){
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
                    if (cell!=null) {
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

    }*/

    /*---------------------------------------------------------------------------------------------------------------------------------------*/
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
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<Map<String, Object>> values = getHeaderFilterValuesNS(sheet, headers, "tipo_cliente", "Nuevo", "Nuevo");

            int currentRow = 0;
            int rowsPerBatch = 5000;

            /*for (Map<String, Object> rowData : values) {
                for (String fields : requiredFields) {

                    for (Map.Entry<String, Object> entry : rowData.entrySet()){
                        if (entry.getKey().contains(fields)) {
                            System.out.println("CAMPO: " + entry.getKey() + " - VALOR: " + entry.getValue());
                        }
                    }

                    currentRow++;

                    if (currentRow % rowsPerBatch == 0) {
                        runtime();
                        Thread.sleep(200);
                    }

                }
                System.out.println();
            }*/

            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, values, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            //headers = getHeaders(sheet);

            //values = getHeaderValuesN(sheet, headers);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

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
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, "Cartera Bruta", fechaCorte);

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


    public static void nuevosOficinas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, "tipo_cliente", "Nuevo", "Nuevo");

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            workbook.close();
            runtime();
            waitSeconds(2);


        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void carteraBruta(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile/*, List<Map<String, String>> datosMasterFile*/) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        System.setProperty("org.apache.poi.ooxml.strict", "false");

        try {
            Workbook workbook = WorkbookFactory.create(new File(okCarteraFile));


            IOUtils.setByteArrayMaxOverride(20000000);

            Sheet sheet = workbook.getSheetAt(0);

            List<String> headers = getHeadersN(sheet);
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            System.out.println("EL ANÁLISIS PUEDE SER ALGO DEMORADO POR FAVOR ESPERE...");

            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            List<Map<String, Object>> datosFiltrados = getHeaderFilterValuesNS(sheet, headers, campoFiltrar, valorInicio, valorFin);

            System.out.println();
            System.out.println("CREANDO ARCHIVO TEMPORAL");
            crearNuevaHojaExcel(camposDeseados, datosFiltrados, tempFile);

            workbook = WorkbookFactory.create(new File(tempFile));

            sheet = workbook.getSheetAt(0);

            System.out.println("SHEET_NAME TEMP_FILE: " + sheet.getSheetName());
            List<String> errores = new ArrayList<>();
            List<String> coincidencias = new ArrayList<>();

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));
            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                String error = hoja + " -> LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                System.out.println(error);
                                errores.add(error);

                            } else {

                                String coincidencia = hoja + " -> LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey();
                                System.out.println(coincidencia);
                                coincidencias.add(coincidencia);

                            }
                        } else {
                            //System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            logWinsToFile(masterFile, coincidencias);
            logErrorsToFile(masterFile, errores);
            workbook.close();
            runtime();
            waitSeconds(2);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        System.setProperty("org.apache.poi.ooxml.strict", "true");
    }

}
