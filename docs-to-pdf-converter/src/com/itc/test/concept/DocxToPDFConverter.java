package com.itc.test.concept;

import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;

public class DocxToPDFConverter extends Converter {

    public DocxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {
        loading();

        XWPFDocument document = new XWPFDocument(inStream);
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();

        // Configurar márgenes
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setTop(BigInteger.valueOf(360));
        pageMar.setBottom(BigInteger.valueOf(360));
        pageMar.setLeft(BigInteger.valueOf(360));
        pageMar.setRight(BigInteger.valueOf(360));

        // Procesar todos los párrafos y tablas
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                configureParagraph((XWPFParagraph) element);
            } else if (element instanceof XWPFTable) {
                configureTable((XWPFTable) element);
            }
        }

        // Configurar página
        CTPageSz pageSize = sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(11906));
        pageSize.setH(BigInteger.valueOf(16838));

        // Configurar opciones PDF
        PdfOptions options = PdfOptions.create();
        options.fontEncoding("Identity-H");
        
        processing();
        
        PdfConverter.getInstance().convert(document, outStream, options);

        finished();
    }

    private void configureParagraph(XWPFParagraph para) {
        CTPPr ppr = para.getCTP().getPPr();
        if (ppr == null) ppr = para.getCTP().addNewPPr();

        // Configurar espaciado
        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setBefore(new BigInteger("60"));
        spacing.setAfter(new BigInteger("60"));
        spacing.setLine(new BigInteger("240"));
        
        // Manejar indentación
        CTInd ind = ppr.isSetInd() ? ppr.getInd() : ppr.addNewInd();
        
        // Configurar indentación base
        ind.setLeft(BigInteger.valueOf(0));
        ind.setFirstLine(BigInteger.valueOf(0));
        
        // Si el párrafo tiene estilo de lista
        if (para.getStyle() != null && para.getStyle().toLowerCase().contains("list")) {
            ind.setLeft(BigInteger.valueOf(360));
            ind.setHanging(BigInteger.valueOf(360));
        }
    }

    private void configureTable(XWPFTable table) {
        // Configurar propiedades de la tabla
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) {
            tblPr = table.getCTTbl().addNewTblPr();
        }

        // Configurar ancho de tabla
        CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        tblWidth.setW(BigInteger.valueOf(9000));
        tblWidth.setType(STTblWidth.DXA);

        // Asegurar que las celdas vacías se muestren
        CTTblLook tblLook = tblPr.isSetTblLook() ? tblPr.getTblLook() : tblPr.addNewTblLook();
        tblLook.setFirstRow(true);
        tblLook.setLastRow(true);
        tblLook.setFirstColumn(true);
        tblLook.setLastColumn(true);
        tblLook.setNoHBand(true);
        tblLook.setNoVBand(true);

        // Procesar cada celda
        for (XWPFTableRow row : table.getRows()) {
            CTTrPr rowPr = row.getCtRow().addNewTrPr();
            rowPr.addNewTrHeight().setVal(BigInteger.valueOf(360)); // Altura mínima de fila

            for (XWPFTableCell cell : row.getTableCells()) {
                // Configurar propiedades de celda
                CTTcPr tcPr = cell.getCTTc().getTcPr();
                if (tcPr == null) {
                    tcPr = cell.getCTTc().addNewTcPr();
                }

                // Asegurar que la celda tenga un ancho
                CTTblWidth cellWidth = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                cellWidth.setType(STTblWidth.DXA);
                cellWidth.setW(BigInteger.valueOf(1800)); // Ancho mínimo de celda

                // Asegurar que las celdas vacías mantengan su espacio
                if (cell.getText().trim().isEmpty()) {
                    cell.setText(" ");
                }

                // Preservar el contenido y formato de los párrafos dentro de la celda
                for (XWPFParagraph p : cell.getParagraphs()) {
                    // Configurar el párrafo dentro de la celda
                    CTPPr ppr = p.getCTP().getPPr();
                    if (ppr == null) {
                        ppr = p.getCTP().addNewPPr();
                    }

                    // Configurar espaciado del párrafo dentro de la celda
                    CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
                    spacing.setBefore(BigInteger.valueOf(0));
                    spacing.setAfter(BigInteger.valueOf(0));
                    spacing.setLine(BigInteger.valueOf(240));

                    // Preservar la alineación del párrafo
                    if (p.getAlignment() != ParagraphAlignment.LEFT) {
                        p.setAlignment(ParagraphAlignment.LEFT);
                    }
                }
            }
        }
    }
}