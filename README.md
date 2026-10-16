# QC Comparative Dissolution Report Generator

This starter is adapted to the uploaded `CDP_f2_Valsartan_10-160mg.docx`.

## Workflow
1. Deploy this repository to Render using the included `render.yaml` Blueprint.
2. Open the Render URL and download the Excel input template.
3. Update the Product sheet and the pH 1.2, pH 4.5, and pH 6.8 sheets.
4. Upload the completed `.xlsx` workbook and click Generate Report.
5. Download the ZIP containing Word and PDF outputs.

The workbook starts with values extracted from the uploaded report as examples. Replace them before generating a new study. The app calculates f2, updates data tables, and redraws the three charts.

## Deploy
Upload all project files to a GitHub repository. In Render choose New + → Blueprint and connect the repository. Render will use `render.yaml`; `apt.txt` requests LibreOffice for PDF conversion.

## Review notes
This is a starter and requires validation before official QC use. The calculation implemented is f2 = 50 × log10(100 / sqrt(1 + mean((Reference − Test)^2))). Check it against your approved QC procedure. PDF pagination can vary with fonts and LibreOffice. The app replaces the native embedded charts with generated chart images and keeps their approximate dimensions. Add authentication and access controls before exposing confidential data publicly.
