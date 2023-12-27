package org.main;

import org.dataTest.Start;

import static org.dataTest.historicoCarteraSegMonto_ColocPorOF.HistoricoCarteraSegMonto_ColocPorOF.*;

public class Main {
    public static void main(String[] args) {

        try {
            //test();
            //testWithJson();
            //testWithNew();

            testWithNewMasterFile();

            //Start.excecution();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
}