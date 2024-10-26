package xmlinvoiceviewer;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableColumn;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class XMLInvoiceViewer {
    private JFrame frame;
    private JTable table;
    private DefaultTableModel tableModel;
    private JComboBox<String> compradorComboBox;
    private List<String> archivosConErrores;

    public XMLInvoiceViewer() {
        // Crear la ventana principal
        frame = new JFrame("Lector de Facturas XML");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLayout(new BorderLayout());

        // Inicializar la lista de archivos con errores
        archivosConErrores = new ArrayList<>();

        // Estilo visual de la ventana principal
        frame.getContentPane().setBackground(new Color(204, 255, 204)); // Fondo verde claro

        // Crear el modelo de tabla con columnas visibles y ocultas
        String[] columnNames = {"Proveedor", "Número de Factura", "Fecha", "Nombre del Item", "Cantidad", "Valor", "IVA", "Valor en Pesos", "IVA en Pesos", "Total en Pesos", "Centro de Costos"};
        tableModel = new DefaultTableModel(columnNames, 0);
        table = new JTable(tableModel);

        // Estilo de la tabla
        table.setBackground(new Color(204, 255, 204)); // Fondo verde claro
        table.setForeground(Color.BLACK); // Texto negro
        table.setFont(new Font("Arial", Font.PLAIN, 14));
        table.getTableHeader().setBackground(new Color(0, 153, 0)); // Encabezado verde oscuro
        table.getTableHeader().setForeground(Color.WHITE); // Texto blanco en encabezado
        table.getTableHeader().setFont(new Font("Arial", Font.BOLD, 16));

        // Configurar la columna "Centro de Costos" para usar JComboBox como editor de celdas
        TableColumn centroCostosColumn = table.getColumnModel().getColumn(10);
        JComboBox<String> comboBox = new JComboBox<>(new String[]{"Administración", "Deposito", "Parqueadero", "SZU-505", "STA-068", "STE-436", "STE-456", "STE-421", "TTZ-648", "WCP-392", "UIC-841", "TNH-287", "SZV-209", "GDX-212"});
        centroCostosColumn.setCellEditor(new DefaultCellEditor(comboBox));

        // Agregar la tabla a un scroll pane
        JScrollPane scrollPane = new JScrollPane(table);
        frame.add(scrollPane, BorderLayout.CENTER);

        // Crear el selector de comprador
        JPanel compradorPanel = new JPanel();
        compradorPanel.setBackground(new Color(204, 255, 204)); // Fondo verde claro
        compradorPanel.setLayout(new FlowLayout(FlowLayout.LEFT));
        compradorPanel.add(new JLabel("Comprador: "));
        compradorComboBox = new JComboBox<>(new String[]{"LEONARDO ANTONIO GONZALEZ CARMONA", "TRANSPORTES Y VOLQUETAS GONZALEZ SAS", "GRUPO NUTABE SAS"}); // Agregar más compradores aquí
        compradorComboBox.setFont(new Font("Arial", Font.BOLD, 14));
        compradorComboBox.setBackground(new Color(204, 255, 204)); // Fondo verde claro
        compradorComboBox.setForeground(Color.BLACK); // Texto negro
        compradorPanel.add(compradorComboBox);
        frame.add(compradorPanel, BorderLayout.NORTH);

        // Crear botones para seleccionar archivo o carpeta
        JPanel buttonPanel = new JPanel();
        buttonPanel.setBackground(new Color(102, 51, 0)); // Fondo marrón oscuro
        JButton selectFileButton = new JButton("Seleccionar Archivo XML");
        JButton selectFolderButton = new JButton("Seleccionar Carpeta");
        JButton exportToExcelButton = new JButton("Exportar a Excel");
        JButton verErroresButton = new JButton("Ver Archivos con Errores");

        // Estilo de los botones
        selectFileButton.setBackground(new Color(255, 140, 0)); // Naranja claro
        selectFileButton.setForeground(Color.BLACK); // Texto negro
        selectFileButton.setFont(new Font("Arial", Font.BOLD, 14));

        selectFolderButton.setBackground(new Color(255, 140, 0)); // Naranja claro
        selectFolderButton.setForeground(Color.BLACK); // Texto negro
        selectFolderButton.setFont(new Font("Arial", Font.BOLD, 14));

        exportToExcelButton.setBackground(new Color(255, 140, 0)); // Naranja claro
        exportToExcelButton.setForeground(Color.WHITE); // Texto blanco
        exportToExcelButton.setFont(new Font("Arial", Font.BOLD, 14));

        verErroresButton.setBackground(new Color(255, 140, 0)); // Naranja claro
        verErroresButton.setForeground(Color.BLACK); // Texto negro
        verErroresButton.setFont(new Font("Arial", Font.BOLD, 14));

        buttonPanel.add(selectFileButton);
        buttonPanel.add(selectFolderButton);
        buttonPanel.add(exportToExcelButton);
        buttonPanel.add(verErroresButton);
        frame.add(buttonPanel, BorderLayout.SOUTH);

        // Acción para seleccionar un archivo XML
        selectFileButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int option = fileChooser.showOpenDialog(frame);
                if (option == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    procesarArchivoXML(selectedFile);
                }
            }
        });

        // Acción para seleccionar una carpeta que contiene archivos XML
        selectFolderButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser folderChooser = new JFileChooser();
                folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int option = folderChooser.showOpenDialog(frame);
                if (option == JFileChooser.APPROVE_OPTION) {
                    File selectedFolder = folderChooser.getSelectedFile();
                    for (File file : selectedFolder.listFiles()) {
                        if (file.getName().endsWith(".xml")) {
                            procesarArchivoXML(file);
                        }
                    }
                }
            }
        });

        // Acción para exportar los datos de la tabla a un archivo Excel
        exportToExcelButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser folderChooser = new JFileChooser();
                folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int option = folderChooser.showSaveDialog(frame);
                if (option == JFileChooser.APPROVE_OPTION) {
                    File selectedFolder = folderChooser.getSelectedFile();
                    exportarTablaAExcel(selectedFolder);
                }
            }
        });

        // Acción para ver los archivos con errores
        verErroresButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (archivosConErrores.isEmpty()) {
                    JOptionPane.showMessageDialog(frame, "No se encontraron errores en los archivos procesados.", "Archivos con Errores", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    StringBuilder errorMessage = new StringBuilder("Archivos con errores:\n");
                    for (String archivo : archivosConErrores) {
                        errorMessage.append(archivo).append("\n");
                    }
                    JOptionPane.showMessageDialog(frame, errorMessage.toString(), "Archivos con Errores", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Mostrar la ventana principal
        frame.setVisible(true);
    }

    private void procesarArchivoXML(File xmlFile) {
        try {
            // Configurar el parser para leer el archivo XML
            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            dbFactory.setNamespaceAware(true); // Hacer que el parser sea consciente de los espacios de nombres
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(xmlFile);
            doc.getDocumentElement().normalize();

            Element rootElement = doc.getDocumentElement();

            // Obtener proveedor, número de factura y fecha
            String proveedor = getOptionalTagValue(rootElement, "cac:AccountingSupplierParty", "cbc:Name");
            String idFactura = getOptionalTagValue(rootElement, "cbc:ID");
            String fecha = getOptionalTagValue(rootElement, "cbc:IssueDate");

            // Obtener los ítems de la factura
            NodeList items = doc.getElementsByTagNameNS("*", "InvoiceLine");
            String comprador = (String) compradorComboBox.getSelectedItem(); // Obtener el comprador seleccionado
            for (int i = 0; i < items.getLength(); i++) {
                Node itemNode = items.item(i);
                if (itemNode.getNodeType() == Node.ELEMENT_NODE) {
                    Element itemElement = (Element) itemNode;
                    String nombreItem = getOptionalTagValue(itemElement, "cbc:Description");
                    String cantidad = getOptionalTagValue(itemElement, "cbc:InvoicedQuantity");
                    String valor = getOptionalTagValue(itemElement, "cbc:LineExtensionAmount");
                    String iva = getOptionalTagValue(itemElement, "cbc:TaxAmount");

                    // Convertir valores a formato pesos (usando puntos como separadores decimales y formatear)
                    double valorDouble = valor.isEmpty() ? 0.0 : Double.parseDouble(valor.replace(",", "."));
                    double ivaDouble = iva.isEmpty() ? 0.0 : Double.parseDouble(iva.replace(",", "."));
                    double totalDouble = valorDouble + ivaDouble;
                    NumberFormat currencyFormat = NumberFormat.getCurrencyInstance(new Locale("es", "CO"));
                    String valorEnPesos = currencyFormat.format(valorDouble);
                    String ivaEnPesos = currencyFormat.format(ivaDouble);
                    String totalEnPesos = currencyFormat.format(totalDouble);

                    // Agregar los datos a la tabla, el centro de costos se seleccionará luego en la tabla
                    tableModel.addRow(new Object[]{proveedor, idFactura, fecha, nombreItem, cantidad, valor, iva, valorEnPesos, ivaEnPesos, totalEnPesos, "Administración"});
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            archivosConErrores.add(xmlFile.getName()); // Agregar archivo a la lista de errores
            JOptionPane.showMessageDialog(frame, "Error al procesar el archivo: " + xmlFile.getName(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    // Método auxiliar para obtener el valor de un tag específico, devolviendo un valor vacío si no existe
    private String getOptionalTagValue(Element parentElement, String parentTagName, String childTagName) {
        NodeList parentList = parentElement.getElementsByTagName(parentTagName);
        if (parentList.getLength() > 0) {
            Element parent = (Element) parentList.item(0);
            return getOptionalTagValue(parent, childTagName);
        }
        return "";
    }

    private String getOptionalTagValue(Element element, String tagName) {
        NodeList nodeList = element.getElementsByTagName(tagName);
        if (nodeList.getLength() > 0) {
            return nodeList.item(0).getTextContent();
        }
        return "";
    }

    private void exportarTablaAExcel(File directory) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Facturas");

        // Crear la fila de encabezado
        Row headerRow = sheet.createRow(0);
        String[] exportColumns = {"Comprador", "Proveedor", "Número de Factura", "Fecha", "Nombre del Item", "Cantidad", "Valor", "IVA", "Total en Pesos", "Centro de Costos"};
        for (int i = 0; i < exportColumns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(exportColumns[i]);
        }

        // Rellenar las filas con los datos de la tabla
        String comprador = (String) compradorComboBox.getSelectedItem(); // Obtener el comprador seleccionado
        for (int i = 0; i < tableModel.getRowCount(); i++) {
            Row row = sheet.createRow(i + 1);
            int colIndex = 0;
            // Agregar la columna de comprador al principio
            Cell compradorCell = row.createCell(colIndex++);
            compradorCell.setCellValue(comprador);
            for (int j = 0; j < tableModel.getColumnCount(); j++) {
                // Exportar solamente las columnas especificadas (excluir "Valor en Pesos" e "IVA en Pesos")
                if (j != 7 && j != 8) { // Excluir columnas "Valor en Pesos" e "IVA en Pesos"
                    Cell cell = row.createCell(colIndex++);
                    Object value = tableModel.getValueAt(i, j);
                    cell.setCellValue(value != null ? value.toString() : "");
                }
            }
        }

        // Guardar el archivo Excel en la carpeta seleccionada
        try {
            FileOutputStream fileOut = new FileOutputStream(new File(directory, "facturas.xlsx"));
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            JOptionPane.showMessageDialog(frame, "Exportado a Excel exitosamente", "Éxito", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(frame, "Error al exportar a Excel", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new XMLInvoiceViewer());
    }
}
