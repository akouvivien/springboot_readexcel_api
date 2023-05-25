package com.gs2e.xlxreadapi.Helper;

import com.gs2e.xlxreadapi.Model.Electeur;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] HEADERs = {"Nom", "Prenom", "Service" };
    static String SHEET = "Electeurs";

    public static boolean hasExcelFormat(MultipartFile file) {

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }
//creation du fichier excel a partir de la liste des electeurs
    public static ByteArrayInputStream electeursToExcel(List<Electeur> electeurs) {

        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream();) {
            Sheet sheet = workbook.createSheet(SHEET);

            // Header
            Row headerRow = sheet.createRow(0);

            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }

            int rowIdx = 1;
            for (Electeur electeur : electeurs) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(electeur.getNom());
                row.createCell(1).setCellValue(electeur.getPrenom());
                row.createCell(2).setCellValue(electeur.getService());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
        }
    }
// Creation de la liste d'electeur à partir du contenu du fichier excel
    public static List<Electeur> excelToElecteurs(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet(SHEET);

            //creation d'un ietrateur
            Iterator<Row> rows = sheet.iterator();

            List<Electeur> electeurs = new ArrayList<Electeur>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Electeur electeur = new Electeur();

                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    switch (cellIdx) {
                        case 0:
                            electeur.setNom(currentCell.getStringCellValue());
                            break;

                        case 1:
                            electeur.setPrenom(currentCell.getStringCellValue());
                            break;

                        case 2:
                            electeur.setService(currentCell.getStringCellValue());
                            break;

                        default:
                            break;
                    }

                    cellIdx++;
                }

                electeurs.add(electeur);
            }

            workbook.close();

            return electeurs;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }
}