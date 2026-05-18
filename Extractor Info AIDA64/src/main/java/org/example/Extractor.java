package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Extractor extends JFrame {
    private JTextField txtRuta;
    private JTable tabla;
    private DefaultTableModel modeloTabla;
    private List<Equipo> listaEquipos = new ArrayList<>();

    // 1. Columna renombrada a "Chapa"
    private final String[] columnas = {
            "Chapa", "Nombre PC", "Marca PC", "Modelo PC", "Nº Serie PC", "MAC", "IP", "Licencia",
            "Monitor ID", "Monitor Nombre", "Monitor Modelo"
    };

    public Extractor() {
        setTitle("Extractor AIDA64 Profesional - Java 11");
        setSize(1350, 700);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JPanel panelNorte = new JPanel(new FlowLayout(FlowLayout.LEFT));
        txtRuta = new JTextField(60);
        JButton btnSeleccionar = new JButton("Seleccionar Carpeta");
        panelNorte.add(new JLabel("Ruta:"));
        panelNorte.add(txtRuta);
        panelNorte.add(btnSeleccionar);

        modeloTabla = new DefaultTableModel(columnas, 0);
        tabla = new JTable(modeloTabla);
        tabla.setAutoCreateRowSorter(true);

        JButton btnExportar = new JButton("Generar Reporte Excel");

        btnExportar.setFont(new java.awt.Font("Arial", java.awt.Font.BOLD, 14));
        btnExportar.setBackground(new java.awt.Color(33, 115, 70));
        btnExportar.setForeground(java.awt.Color.WHITE);

        add(panelNorte, BorderLayout.NORTH);
        add(new JScrollPane(tabla), BorderLayout.CENTER);
        add(btnExportar, BorderLayout.SOUTH);

        btnSeleccionar.addActionListener(e -> seleccionarCarpeta());
        btnExportar.addActionListener(e -> exportarExcel());

        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void seleccionarCarpeta() {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File carpeta = chooser.getSelectedFile();
            txtRuta.setText(carpeta.getAbsolutePath());
            procesarArchivos(carpeta);
        }
    }

    private void procesarArchivos(File carpeta) {
        File[] archivos = carpeta.listFiles((dir, name) -> name.toLowerCase().endsWith(".htm"));
        if (archivos == null || archivos.length == 0) {
            JOptionPane.showMessageDialog(this, "No se encontraron archivos .htm");
            return;
        }

        listaEquipos.clear();
        modeloTabla.setRowCount(0);

        for (File f : archivos) {
            try {
                Equipo eq = extraerInfo(f);
                listaEquipos.add(eq);
                modeloTabla.addRow(eq.toObjectArray());
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private Equipo extraerInfo(File archivo) throws IOException {
        Document doc = Jsoup.parse(archivo, "ISO-8859-1");
        Equipo eq = new Equipo();

        // 2. Lógica para quitar el ".htm" del nombre
        String nombreArchivo = archivo.getName();
        if (nombreArchivo.toLowerCase().endsWith(".htm")) {
            eq.chapa = nombreArchivo.substring(0, nombreArchivo.length() - 4);
        } else {
            eq.chapa = nombreArchivo;
        }

        Elements filas = doc.select("tr");

        for (Element fila : filas) {
            String etiqueta = obtenerEtiqueta(fila);
            String valor = obtenerValor(fila);

            if (valor.isEmpty()) continue;

            switch (etiqueta) {
                case "Computer Name": eq.nombre = valor; break;
                case "DMI System Manufacturer": eq.marca = valor; break;
                case "DMI System Product": eq.modelo = valor; break;
                case "DMI System Serial Number": eq.serie = valor; break;
                case "Primary MAC Address": eq.mac = valor; break;
                case "Primary IP Address": eq.ip = valor; break;
                case "Product Key": eq.licencia = valor; break;
                case "Monitor ID": eq.monitorId = valor; break;
                case "Monitor Name": eq.monitorNombre = valor; break;
                case "Model":
                case "Monitor Model":
                case "Monitor Vendor":
                    eq.monitorModelo = valor;
                    break;
            }
        }
        return eq;
    }

    private String obtenerEtiqueta(Element fila) {
        Elements celdas = fila.select("td");
        if (celdas.size() >= 2) {
            for (int i = celdas.size() - 2; i >= 0; i--) {
                String texto = celdas.get(i).text().trim();
                if (!texto.isEmpty()) return texto;
            }
        }
        return "";
    }

    private String obtenerValor(Element fila) {
        Elements celdas = fila.select("td");
        if (celdas.size() >= 2) {
            String val = celdas.last().text().trim();
            return val.replace("[NoDB]", "").trim();
        }
        return "";
    }

    private void exportarExcel() {
        if (listaEquipos.isEmpty()) return;

        JFileChooser saver = new JFileChooser();
        if (saver.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            String path = saver.getSelectedFile().getAbsolutePath();
            if (!path.toLowerCase().endsWith(".xlsx")) path += ".xlsx";

            try (Workbook wb = new XSSFWorkbook(); FileOutputStream fos = new FileOutputStream(path)) {
                Sheet sheet = wb.createSheet("Inventario");
                Row header = sheet.createRow(0);
                CellStyle style = wb.createCellStyle();

                org.apache.poi.ss.usermodel.Font poiFont = wb.createFont();
                poiFont.setBold(true);
                style.setFont(poiFont);

                for (int i = 0; i < columnas.length; i++) {
                    Cell cell = header.createCell(i);
                    cell.setCellValue(columnas[i]);
                    cell.setCellStyle(style);
                }

                for (int i = 0; i < listaEquipos.size(); i++) {
                    Row row = sheet.createRow(i + 1);
                    Object[] data = listaEquipos.get(i).toObjectArray();
                    for (int j = 0; j < data.length; j++) {
                        row.createCell(j).setCellValue(data[j].toString());
                    }
                }

                for (int i = 0; i < columnas.length; i++) sheet.autoSizeColumn(i);
                wb.write(fos);
                JOptionPane.showMessageDialog(this, "Excel generado con éxito.");
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
    }

    private static class Equipo {
        String chapa="", nombre="", marca="", modelo="", serie="", mac="", ip="", licencia="";
        String monitorId="", monitorNombre="", monitorModelo="";

        Object[] toObjectArray() {
            return new Object[]{
                    chapa, nombre, marca, modelo, serie, mac, ip, licencia,
                    monitorId, monitorNombre, monitorModelo
            };
        }
    }
}