/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import javax.swing.JFileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileOutputStream;

/**
 *
 * @author khadijahela
 */
public class Excel {

    public static void main(String[] args) throws IOException {

        Excel.writeExcel(Excel.readExcel());

    }

    public static List<Incident> readExcel() throws FileNotFoundException, IOException {
        JFileChooser jf = new JFileChooser();
        jf.showDialog(null, "choisir un fichier");
        File myFile = jf.getSelectedFile();

        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);

        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
        // 29 -> date
        // 1 -> incident
        // 30 -> operateur
        Date date = null;
        String id = null;
        String operateur = null;

        List<Incident> listeIncidents = new ArrayList<>();

        rowIterator.next();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            //For each row, iterate through each columns 

            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getColumnIndex()) {
                    case 1:
                        id = (cell.getStringCellValue());
                        break;
                    case 29:
                        date = (cell.getDateCellValue());
                        break;
                    case 30:
                        operateur = (cell.getStringCellValue());
                        break;
                    default:
                        break;
                }

            }

            Incident incident = new Incident(id, date, operateur);
            listeIncidents.add(incident);

        }
        System.out.println("lecture terminée ");

        List<Incident> maliste = Excel.retournerOperateurParIncident(listeIncidents);

        return maliste;
    }

    public static List<Incident> retournerOperateurParIncident(List<Incident> listeIncidents) {
        List<Incident> maListe = new ArrayList<>();//operateur par incident
        List<Incident> maListecourante = new ArrayList<>();//liste des operateur de meme id

        for (int i = 0; i < listeIncidents.size(); i++) {
            Incident incident = listeIncidents.get(i);
            int dernpos = compter(listeIncidents, incident);
            // System.out.println(incident.getId()+" : "+dernpos);

            for (int j = i; j <= i + dernpos; j++) {
                maListecourante.add(listeIncidents.get(j)); // ajouter les incident de meme id en malistecourante
            }
            if (dernpos + 1 < listeIncidents.size()) {
                i = i + dernpos;
            } else {
                i = listeIncidents.size();
            }
            maListe.add(retourOperateur(maListecourante));
            maListecourante = new ArrayList<>();

        }

        return maListe;
    }

    public static Incident retourOperateur(List<Incident> liste) {
        for (Incident i : liste) {
            if (!i.getOperateur().equals("temip.sm")) {
                return i;//retourne nom du premier operateur 
            }
        }
        return liste.get(0);//sinon temip.sm

    }

    public static int compter(List<Incident> listeIncidents, Incident incident) {
        int ee = 0;
        for (int i = 0; i < listeIncidents.size(); i++) {
            if (listeIncidents.get(i).getId().equals(incident.getId())) {
                ee++;//nombre incident de même id
            }
        }
        return ee;
    }

//    public static List<String> retournerListeIncident(List<Incident> listeIncidents) {
//        List<String> incident = new ArrayList<>();
//        for (int i = 1; i < listeIncidents.size(); i++) {
//            if (!listeIncidents.get(i).getId().equals(listeIncidents.get(i - 1).getId())) {
//                incident.add(listeIncidents.get(i).getId());
//            }
//        }
//        return incident;
//    }
    public static void writeExcel(List<Incident> listeIncidents) {
        //create a new workbook
        Workbook wb = new XSSFWorkbook();

        //add a new sheet to the workbook
        Sheet sheet1 = wb.createSheet("Resultat");

        Row row = sheet1.createRow(0);

        Cell row1col1 = row.createCell(0);
        Cell row1col2 = row.createCell(1);
        Cell row1col3 = row.createCell(2);

        sheet1.setColumnWidth(0, 5000);
        sheet1.setColumnWidth(1, 5000);
        sheet1.setColumnWidth(2, 7000);

        row1col1.setCellValue("INCIDENT");
        row1col2.setCellValue("DATE");
        row1col3.setCellValue("OPERATEUR");

        for (int i = 0; i < listeIncidents.size(); i++) {

            row = sheet1.createRow(i + 2);

            row1col1 = row.createCell(0);
            row1col2 = row.createCell(1);
            row1col3 = row.createCell(2);

            //incident
            row1col1.setCellValue(listeIncidents.get(i).getId());

            //date
            CellStyle cellStyle = wb.createCellStyle();
            CreationHelper createHelper = wb.getCreationHelper();
            short dateFormat = createHelper.createDataFormat().getFormat("yyyy-dd-MM hh:mm");
            cellStyle.setDataFormat(dateFormat);

            row1col2.setCellStyle(cellStyle);
            row1col2.setCellValue(listeIncidents.get(i).getDate());

            //operateur
            row1col3.setCellValue("  "+listeIncidents.get(i).getOperateur());

        }

        //write the excel to a file
        JFileChooser jf = new JFileChooser();
        jf.showSaveDialog(null);
        try {
            FileOutputStream fileOut = new FileOutputStream(jf.getSelectedFile() + ".xlsx");
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("écriture terminée");

    }

}
