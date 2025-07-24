// InDesign ExtendScript: ლათინური განლაგების ქართული ტექსტის უნიკოდში გადაყვანა
// ვერსია: 3.7 - დამატებულია მრავალჯერადი მონიშნული ობიექტის დამუშავების მხარდაჭერა
// ავტორი: გიორგი (საწყისი იდეა და მრავალი ტესტირება) და Gemini (დახმარება და განვითარება)
// თარიღი: 2025-05-23

// --- სიმბოლოების რუკები ---
var charmaps = {
    utf: "აბგდევზთიკლმნოპჟრსტუფქღყშჩცძწჭხჯჰ“„", 
    lat: "abgdevzTiklmnopJrstufqRySCcZwWxjh~`"  
};

// --- ტექსტის კონვერტაციის ფუნქცია ---
function convertText(originalText) {
    var convertedText = "";
    var fromChars = charmaps.lat;
    var toChars = charmaps.utf;
    var conversionHappened = false;
    var convertedCharCount = 0; 

    for (var i = 0; i < originalText.length; i++) {
        var charToConvert = originalText.charAt(i);
        var index = fromChars.indexOf(charToConvert); 

        if (index !== -1 && index < toChars.length) {
            var newChar = toChars.charAt(index);
            convertedText += newChar;
            if (charToConvert !== newChar) { 
                 conversionHappened = true;
                 convertedCharCount++; 
            }
        } else {
            convertedText += charToConvert;
        }
    }
    return { text: convertedText, changed: conversionHappened, count: convertedCharCount };
}

// --- დამხმარე ფუნქცია დაბლოკვის შესამოწმებლად ---
function isObjectLocked(item) {
    try {
        if (item && item.isValid) { 
            if (item.parentTextFrames && item.parentTextFrames.length > 0) {
                var parentTF = item.parentTextFrames[0];
                if (parentTF instanceof TextFrame && parentTF.isValid) {
                    try { if (parentTF.locked) return true; } catch (e_tf_lock) { /* იგნორირება */ }
                    if (parentTF.itemLayer instanceof Layer && parentTF.itemLayer.isValid) {
                        try { if (parentTF.itemLayer.locked) return true; } catch (e_layer_lock) { /* იგნორირება */ }
                    }
                }
            }
            if (item.parentStory instanceof Story && item.parentStory.isValid) {
                try { if (item.parentStory.locked) return true; } catch (e_story_lock) { /* იგნორირება */ }
            }
        }
    } catch (e) { /* იგნორირება */ }
    return false; 
}


// --- მთავარი ფუნქცია ---
function runScopedConversion() {
    if (app.documents.length === 0) {
        alert("გთხოვთ, გახსნათ დოკუმენტი.");
        return;
    }
    var doc = app.activeDocument;

    var mainDialog = app.dialogs.add({ name: "ტექსტის კოდირების გადაყვანა" });
    var dialogColumn = mainDialog.dialogColumns.add();
    
    var scopeLabelRow = dialogColumn.dialogRows.add(); 
    scopeLabelRow.staticTexts.add({staticLabel: "აირჩიეთ მოქმედების არეალი:"});
    
    var radioGroupRow = dialogColumn.dialogRows.add(); 
    var scopeRadioGroup = radioGroupRow.radiobuttonGroups.add();
    scopeRadioGroup.radiobuttonControls.add({staticLabel: "მონიშნული ობიექტ(ებ)ი", checkedState: true}); // შევცვალე ტექსტი
    scopeRadioGroup.radiobuttonControls.add({staticLabel: "მთლიანი Story (მონიშნულის მიხედვით)"});
    scopeRadioGroup.radiobuttonControls.add({staticLabel: "მთლიანი დოკუმენტი"});

    if (!mainDialog.show()) {
        if (mainDialog.isValid) mainDialog.destroy();
        return; 
    }

    var actionType = scopeRadioGroup.selectedButton; 
        
    if (mainDialog.isValid) mainDialog.destroy();

    var objectsToProcess = []; 
    var overallConversionPerformed = false;
    var totalCharsConverted = 0; 
    var nothingToProcessMessage = "";

    switch (actionType) {
        case 0: // მონიშნული ობიექტ(ებ)ი
            if (app.selection.length > 0) {
                for (var sIdx = 0; sIdx < app.selection.length; sIdx++) { // ვივლით ყველა მონიშნულ ელემენტზე
                    var sel = app.selection[sIdx];
                    if (sel instanceof Text) { 
                        objectsToProcess.push(sel); 
                    } else if (sel instanceof TextFrame && sel.texts.length > 0 && sel.texts[0].isValid) { 
                        objectsToProcess.push(sel.texts[0]); 
                    } else if (sel.hasOwnProperty("texts") && sel.texts.length > 0 && sel.texts[0].isValid) { // მაგ. Cell
                        objectsToProcess.push(sel.texts[0]); 
                    } else if (sel instanceof Table) {
                        var cells = sel.cells.everyItem().getElements();
                        for (var c = 0; c < cells.length; c++) {
                            if (cells[c].texts.length > 0 && cells[c].texts[0].isValid) { 
                                objectsToProcess.push(cells[c].texts[0]); 
                            }
                        }
                    }
                }
                if (objectsToProcess.length === 0) { 
                    nothingToProcessMessage = "მონიშნულ ობიექტ(ებ)ში ვერ მოიძებნა დასამუშავებელი ტექსტი ან მხარდაჭერილი ტიპები."; 
                }
            } else {
                nothingToProcessMessage = "გთხოვთ, მონიშნოთ ტექსტი, ტექსტური ჩარჩო(ები), უჯრ(ებ)ი ან ცხრილ(ებ)ი.";
            }
            break;
        case 1: 
            if (app.selection.length > 0 && app.selection[0].hasOwnProperty("parentStory")) { // Story-სთვის ვიღებთ პირველი მონიშნულის Story-ს
                try {
                    var story = app.selection[0].parentStory;
                    if (story.isValid && story.texts.length > 0 && story.texts[0].isValid) { objectsToProcess.push(story.texts[0]); }
                    else { nothingToProcessMessage = "მონიშნული ობიექტის Story ცარიელია ან არასწორია."; }
                } catch (e) { nothingToProcessMessage = "Story-ს მიღება ვერ მოხერხდა: " + e.message; }
            } else {
                nothingToProcessMessage = "Story-ზე ოპერაციისთვის გთხოვთ, მონიშნოთ ობიექტი, რომელიც Story-ს ნაწილია.";
            }
            break;
        case 2: 
            if (doc.stories.length > 0) {
                for (var s = 0; s < doc.stories.length; s++) {
                    var currentStory = doc.stories[s];
                    if (currentStory.isValid && currentStory.texts.length > 0 && currentStory.texts[0].isValid) {
                        objectsToProcess.push(currentStory.texts[0]);
                    }
                }
            }
            if (objectsToProcess.length === 0) { nothingToProcessMessage = "დოკუმენტში ტექსტი ვერ მოიძებნა დასამუშავებლად."; }
            break;
    }

    if (nothingToProcessMessage) {
        alert(nothingToProcessMessage);
        return;
    }
    if (objectsToProcess.length === 0) { 
        alert("დასამუშავებელი ტექსტური ობიექტები ვერ მოიძებნა.");
        return;
    }

    for (var i = 0; i < objectsToProcess.length; i++) {
        var textObject = objectsToProcess[i]; 
        if (!textObject || !textObject.isValid) continue;
        
        // ვამოწმებთ, არის თუ არა ეს textObject პირდაპირი Text მონიშვნა, რომელიც არ არის მთლიანი TextFrame-ის შიგთავსი.
        // ეს ლოგიკა დახვეწას საჭიროებს, რადგან objectsToProcess უკვე შეიცავს Text ობიექტებს.
        // მთავარია, რომ თუ sel იყო Text, მაშინ textObject არის ეს sel.
        var processAsDirectText = false;
        if (actionType === 0 && app.selection.length === 1 && app.selection[0] === textObject && textObject instanceof Text && !(textObject.parent instanceof TextFrame && textObject.parent.texts[0] === textObject) ) {
            // თუ "მონიშნული ობიექტია" არჩეული, მხოლოდ ერთი რამაა მონიშნული, ეს არის ზუსტად ის ობიექტი, რომელსაც ვამუშავებთ,
            // და ეს არის Text ობიექტი, რომელიც არ წარმოადგენს მთლიანი TextFrame-ის შიგთავსს (ანუ, ნაწილობრივი მონიშვნაა)
             processAsDirectText = true;
        }


        if (processAsDirectText) { 
            var originalContents;
            if (isObjectLocked(textObject)) { continue; }
            try { originalContents = String(textObject.contents); } catch(e) { continue; }

            if (originalContents.replace(/\s/g, '').length > 0) {
                var conversionResult = convertText(originalContents);
                if (conversionResult.changed) {
                    overallConversionPerformed = true;
                    totalCharsConverted += conversionResult.count; 
                    try {
                        textObject.contents = conversionResult.text;
                    } catch (errSetting) { /* იგნორირება */ }
                }
            }
        }
        else if (textObject.hasOwnProperty("paragraphs") && textObject.paragraphs.length > 0) {
            var paragraphs = textObject.paragraphs.everyItem().getElements();
            for (var p = 0; p < paragraphs.length; p++) {
                var currentParagraph = paragraphs[p]; 
                if (!currentParagraph.isValid) continue;
                var originalParaText;
                
                if (isObjectLocked(currentParagraph)) { 
                    continue;
                }
                
                try {
                    originalParaText = String(currentParagraph.contents);
                } catch (e) { 
                    continue; 
                }

                if (originalParaText.replace(/\s/g, '').length === 0) {
                    continue;
                }
                
                var conversionResult = convertText(originalParaText); 

                if (conversionResult.changed) {
                    overallConversionPerformed = true; 
                    totalCharsConverted += conversionResult.count; 
                    try {
                        currentParagraph.contents = conversionResult.text; 
                    } catch (errSetting) { /* იგნორირება */ }
                }
            }
        }
    }

    if (overallConversionPerformed) {
        alert("ტექსტი წარმატებით გადაყვანილია (Lat -> UTF). დაკონვერტირდა " + totalCharsConverted + " სიმბოლო.");
    } else {
        if (objectsToProcess.length > 0) { 
             alert("კონვერტაცია არ შესრულდა. შესაძლოა, ტექსტი არ შეიცავს ლათინურ სიმბოლოებს (როგორც ეს `charmaps.lat`-შია განსაზღვრული) ან უკვე უნიკოდშია.");
        }
    }
}

// --- სკრიპტის გაშვება ---
try {
    app.doScript(runScopedConversion, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "ქართული ტექსტის კონვერტაცია (Lat-UTF)");
} catch(e) {
    alert("მოხდა გაუთვალისწინებელი შეცდომა სკრიპტის მუშაობისას: " + e.message + "\nხაზი: " + e.line);
}

