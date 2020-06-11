package com.rizkykhapidsyah;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileWriter;

public class FormUtama extends JFrame {

    JButton btnExport;
    JTable table;
    JScrollPane scroll;
    DefaultTableModel model;

    String header[] = {"Negara", "Ibukota", "Benua"};
    Object data[][] = {
            {"Indonesia", "Jakarta", "Asia"},
            {"Jepang", "Tokyo", "Asia"},
            {"Inggris", "London", "Eropa"},
            {"Mesir", "Cairo", "Afrika"},
            {"Amerika Serikat", "Washington DC", "Amerika"}
    };

    public FormUtama() {
        super("Export Data Ke Excel by Rizky Khapidsyah");
        Inisialisasi_Komponen();
    }

    private void Inisialisasi_Komponen() {
        model = new DefaultTableModel(data, header);
        table = new JTable();

        table.setModel(model);

        scroll = new JScrollPane(table);
        scroll.setPreferredSize(new Dimension(400, 200));

        add(scroll, BorderLayout.CENTER);

        btnExport = new JButton("Export Data Ke Excel");
        btnExport.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                ExportDataKeExcel(table, new File("DataHasilKonversi.xls"));
            }
        });

        add(btnExport, BorderLayout.SOUTH);

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        pack();
        setVisible(true);
    }

    public void ExportDataKeExcel(JTable table, File file) {
        try {
            TableModel tableModel = table.getModel();
            FileWriter fOut = new FileWriter(file);

            for (int i = 0; i < tableModel.getColumnCount(); i++) {
                fOut.write(tableModel.getColumnName(i) + "\t");
            }

            fOut.write("\n");

            for (int i = 0; i < tableModel.getRowCount(); i++) {
                for (int j = 0; j < tableModel.getColumnCount(); j++) {
                    fOut.write(tableModel.getValueAt(i, j).toString() + "\t");
                }
                fOut.write("\n");
            }

            fOut.close();
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println(e);
        }
    }

}
