/*
 * Export Selection to a New File (PDF/EPS) - Single Object Version
 * Copyright (c) 2025 Gemini
 *
 * აღწერა:
 * ეს სკრიპტი ქმნის ახალ InDesign დოკუმენტს ერთი მონიშნული ობიექტის
 * ზომის მიხედვით, აკოპირებს იქ ობიექტს და შემდეგ ახორციელებს
 * ექსპორტს PDF ან EPS ფორმატში.
 *
 * ვერსია: 5.2 (განახლებული)
 * - კოდი მუშაობს მხოლოდ ერთ მონიშნულ ობიექტზე.
 * - ამოღებულია დაჯგუფების და მრავალჯერადი მონიშვნის ლოგიკა.
 * - Control Panel-დან width და height მნიშვნელობების გამოყენება.
 */

#target indesign;

// მთავარი ფუნქცია
function exportSingleObject() {
    // 1. შემოწმება: უნდა იყოს მონიშნული ზუსტად ერთი ობიექტი
    if (app.documents.length === 0 || app.selection.length !== 1) {
        alert("გთხოვთ, მონიშნეთ მხოლოდ ერთი ობიექტი.");
        return;
    }

    var originalDoc = app.activeDocument;
    var selectedItem = app.selection[0];

    // 2. დიალოგური ფანჯარა ფორმატის ასარჩევად
    var dialog = app.dialogs.add({ name: "Export Selection As" });
    var formatChoice;
    with (dialog.dialogColumns.add()) {
        with (borderPanels.add()) {
            staticTexts.add({ staticLabel: "აირჩიეთ ექსპორტის ფორმატი:" });
            formatChoice = radiobuttonGroups.add();
            with (formatChoice) {
                radiobuttonControls.add({ staticLabel: "PDF", checkedState: true });
                radiobuttonControls.add({ staticLabel: "EPS" });
            }
        }
    }

    // 3. მომხმარებლის არჩევანის დამუშავება
    if (dialog.show() === true) {
        var isPDF = (formatChoice.selectedButton === 0);
        var fileFilter = isPDF ? "Adobe PDF:*.pdf" : "EPS File:*.eps";
        var targetFile = File.saveDialog("შეინახეთ ფაილი როგორც...", fileFilter);
        var tempDoc;

        if (targetFile) {
            try {
                // 4. ობიექტის ზომების აღება Control Panel-დან
                var itemBounds = selectedItem.geometricBounds; // ობიექტის გეომეტრიული საზღვრები
                var itemWidth = itemBounds[3] - itemBounds[1]; // სიგანე
                var itemHeight = itemBounds[2] - itemBounds[0]; // სიმაღლე

                // 5. ახალი დოკუმენტის შექმნა
                tempDoc = app.documents.add();
                tempDoc.documentPreferences.pageWidth = itemWidth;
                tempDoc.documentPreferences.pageHeight = itemHeight;

                with(tempDoc.pages[0].marginPreferences){
                    top = 0; bottom = 0; left = 0; right = 0;
                }
                
                // 6. ობიექტის კოპირება და ჩასმა
                app.activeDocument = originalDoc;
                app.copy(selectedItem);
                
                app.activeDocument = tempDoc;
                app.paste();
                
                var pastedItem = tempDoc.selection[0];
                if (pastedItem) {
                    // ობიექტის განთავსება (0,0) კოორდინატზე
                    var pastedBounds = pastedItem.visibleBounds;
                    var offsetX = -pastedBounds[1];
                    var offsetY = -pastedBounds[0];
                    pastedItem.move(undefined, [offsetX, offsetY]);
                } else {
                    throw new Error("ობიექტის ახალ დოკუმენტში ჩასმა ვერ მოხერხდა.");
                }

                // 7. ექსპორტი
                if (isPDF) {
                    var pdfPreset = app.pdfExportPresets.item("PDF/X-4:2008");
                    if (!pdfPreset.isValid) {
                        try {
                            // PDF/X-4:2008 პრესეტის შექმნა
                            pdfPreset = app.pdfExportPresets.add();
                            pdfPreset.name = "PDF/X-4:2008";
                            pdfPreset.compatibility = PDFCompatibility.ACROBAT_7;
                            pdfPreset.exportReaderSpreads = false;
                            pdfPreset.viewPDF = false;
                            pdfPreset.colorBitmapCompression = BitmapCompression.JPEG;
                            pdfPreset.colorBitmapQuality = PDFCompressionQuality.MAXIMUM;
                            pdfPreset.monochromeBitmapCompression = BitmapCompression.JPEG;
                            pdfPreset.monochromeBitmapQuality = PDFCompressionQuality.MAXIMUM;
                            pdfPreset.transparencyFlattenerPreset = app.transparencyFlattenerPresets.item("[High Resolution]");
                            alert('PDF/X-4:2008 პრესეტი ვერ მოიძებნა და ახალი პრესეტი შეიქმნა.');
                        } catch (e) {
                            alert('PDF/X-4:2008 პრესეტის შექმნა ვერ მოხერხდა: ' + e.message);
                        }
                    }
                    if (!pdfPreset.isValid) {
                        throw new Error('ვერც "PDF/X-4:2008" და ვერც "[High Quality Print]" პრესეტი ვერ მოიძებნა.');
                    }
                    
                    app.pdfExportPreferences.viewPDF = false;
                    tempDoc.exportFile(ExportFormat.PDF_TYPE, targetFile, false, pdfPreset);
                } else {
                    tempDoc.exportFile(ExportFormat.EPS_TYPE, targetFile, false);
                }

                alert("ექსპორტი წარმატებით დასრულდა!");

            } catch (e) {
                alert("ოპერაციისას მოხდა შეცდომა: " + e.message);
            } finally {
                // 8. გასუფთავება
                if (tempDoc && tempDoc.isValid) {
                    tempDoc.close(SaveOptions.NO);
                }
                if (originalDoc && originalDoc.isValid) {
                    app.activeDocument = originalDoc;
                }
            }
        }
        dialog.destroy();
    } else {
        dialog.destroy();
    }
}

// სკრიპტის გაშვება
exportSingleObject();
