// From Erik's Video
// https://www.youtube.com/watch?v=1H4aT4AXY0c

// kauffmann
// https://www.kauffmann.nl/2020/01/18/how-to-use-the-excel-buffer-in-business-central-cloud/

codeunit 50102 "GT - Excel Tools"
{
    /// <summary>
    /// Takes the column # as an integer
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetText(var Buffer: Record "Excel Buffer" temporary; Col: Integer; Row: Integer): Text
    Begin
        if Buffer.Get(Row, Col) then
            exit(Buffer."cell value as text");
    End;
    /// <summary>
    /// Takes the column # as an integer
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetDate(var Buffer: Record "Excel Buffer" temporary; Col: Integer; Row: Integer): Date
    var
        d: Date;
    Begin
        if Buffer.Get(Row, Col) then begin
            Evaluate(d, Buffer."Cell Value as Text");//Evaluate(d, Buffer."Cell Value as Text",9); //ISO8609 did not work in erik's video
            exit(d);
        end;
    End;
    /// <summary>
    /// Takes the column # as an integer
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetDecimal(var Buffer: Record "Excel Buffer" temporary; Col: Integer; Row: Integer): Decimal
    var
        d: Decimal;
    Begin
        if Buffer.Get(Row, Col) then begin
            Evaluate(d, Buffer."Cell Value as Text");
            exit(d);
        end
    End;

    /// <summary>
    /// Takes the column # as an integer
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetInteger(var Buffer: Record "Excel Buffer" temporary; Col: Integer; Row: Integer): Integer
    var
        d: Integer;
    Begin
        if Buffer.Get(Row, Col) then begin
            Evaluate(d, Buffer."Cell Value as Text");
            exit(d);
        end
    End;

    /// <summary>
    /// Takes the column # as a Letter
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetText(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Text
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then
            exit(Buffer."Cell Value as Text");
    end;

    /// <summary>
    /// Takes the column # as a Letter
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetDate(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Date
    var
        d: Date;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(D, Buffer."Cell Value as Text");
            exit(D);
        end;
    end;
    /// <summary>
    /// Takes the column # as a Letter
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetDecimal(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Decimal
    var
        d: Decimal;
    begin
        // Buffer.reset();
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            if Buffer."Cell Value as Text".trim() <> '' then
                Evaluate(d, Buffer."Cell Value as Text".trim())
            else
                d := 0;
            exit(d);
        end;
    end;
    /// <summary>
    /// Takes the column # as a Letter
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetInteger(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Integer
    var
        d: Integer;
    begin
        // Buffer.reset();
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            if Buffer."Cell Value as Text".trim() <> '' then
                Evaluate(d, Buffer."Cell Value as Text".trim())
            else
                d := 0;
            exit(d);
        end;
    end;

    /// <summary>
    /// Takes the column # as an Boolean
    /// </summary>
    /// <param name="Buffer"></param>
    /// <param name="Col"></param>
    /// <param name="Row"></param>
    /// <returns></returns>
    procedure GetBoolean(var Buffer: Record "Excel Buffer" temporary; Col: Integer; Row: Integer): Boolean
    var
        d: Decimal;
    Begin
        if Buffer.Get(Row, Col) then begin
            if Buffer."Cell Value as Text".ToLower() in ['yes', 'true', '1'] then
                exit(true)
            else
                exit(false);
        end;
    End;

    procedure AddTextColumn(var excel: Record "Excel Buffer" temporary; value: Variant)
    begin
        excel.AddColumn(value, false, '', false, false, false, '', excel."Cell Type"::Text);
    end;

    procedure AddNumberColumn(var excel: Record "Excel Buffer" temporary; value: Variant)
    begin
        excel.AddColumn(value, false, '', false, false, false, '', excel."Cell Type"::Number);
    end;

    procedure AddDateColumn(var excel: Record "Excel Buffer" temporary; value: Variant)
    begin
        excel.AddColumn(value, false, '', false, false, false, '', excel."Cell Type"::Date);
    end;

    procedure addFormulaColumn(var excel: Record "Excel Buffer" temporary; Formula: Text)
    begin
        excel.AddColumn(Formula, true, '', false, false, false, '', excel."Cell Type"::Number);
    end;

    procedure SetCell(var Buffer: Record "Excel Buffer" temporary; Col: Integer; Row: Integer; valueToInput: Text)
    var
    begin
        Buffer.SetCurrent(Row, Col);
        Buffer."Cell Value as Text" := valueToInput;
    end;

    procedure GetColumnNumber(ColumnName: Text): Integer
    var
        columnIndex: Integer;
        factor: Integer;
        pos: Integer;
    begin
        factor := 1;
        for pos := strlen(ColumnName) downto 1 do
            if ColumnName[pos] >= 65 then begin
                columnIndex += factor * ((ColumnName[pos] - 65) + 1);
                factor *= 26;
            end;

        exit(columnIndex);
    end;


    
    procedure ImportPhysicalInventoryJournal()
    var
        itemJournal: Record "Item Journal Line";
        excelBuffer: record "Excel Buffer" temporary;
        instr: InStream;
        lastrow: Integer;
        row: Integer;
        FromFile: Text[100];
    begin
        if UploadIntoStream('Import Physical Inventory', '', '', FromFile, instr) then begin
            excelBuffer.OpenBookStream(instr, 'Physical Inventory Journal');
            excelBuffer.ReadSheet();
            excelBuffer.setrange("Column No.", 3); //item no.
            excelBuffer.FindLast();

            LastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            // Skip Headers
            for row := 2 to LastRow do begin
                clear(itemJournal);
                itemJournal.setrange("Posting Date", GetDate(excelBuffer, 'A', row));
                itemJournal.setrange("Document No.", GetText(excelBuffer, 'B', row));
                itemJournal.setrange("Item No.", GetText(excelBuffer, 'C', row));
                itemJournal.setrange("Qty. (Calculated)", GetDecimal(excelBuffer, 'F', row));

                if itemJournal.Findfirst() then begin // we have a match
                    itemJournal.Validate("Qty. (Phys. Inventory)", GetDecimal(excelBuffer, 'H', row));
                    itemJournal.Modify();
                end;
            end;
        end;
    end;

    /// <summary>
    /// Used by the Import Ship-To Addresses
    /// </summary>
    local procedure CreateCustomerShipTo(No: Code[20]; shiptoname: Text; address1: Text; address2: Text; address3: Text; address4: Text; city: Text; state: Text; zip: Text; country: Text)
    var
        shipTo: record "Ship-to Address";
        existingshipto: record "Ship-to Address";
        nextNumber: Integer;
        th: codeunit "Type Helper";
        addressUtility: codeunit "GT - Address Utility";
    begin
        shipTo.init();
        existingshipto.setrange("Customer No.", No);
        if existingshipto.FindLast() then begin
            //this gets the next ship-to number to use
            Evaluate(nextNumber, format(existingshipto.Code));
            nextNumber := nextNumber + 1;
            //Set the Code and Pad it to the next number...
            shipTo.Code := format(nextNumber).PadLeft(5, '0');
        end else begin
            shipTo.code := '00001';
        end;
        shipTo.validate("Customer No.", no);

        shipto.validate(Name, shiptoname);
        shipto.Insert();

        shipTo.Address := address1;
        shipTo."Address 2" := address2;


        //Address
        if address1 = shiptoname.trim() then begin
            // we dont want address1 from QB
            shipTo.validate(Address, address1.trim);
        end
        else begin
            shipTo.validate(Address, address1.trim + ' ' + address2.trim());
        end;

        if strlen(address3.trim() + ' ' + address4.trim()) > 50 then begin
            shipTo.validate("GT QB Extra Data", 'Address 3 & 4: ' + address3.trim() + ' ' + address4.trim() + th.NewLine());
            //copying the first 50 characters
            shipTo.validate("Address 2", CopyStr(format(address3.trim() + ' ' + address4.trim()).trim(), 1, 50));
        end
        else begin
            // it will fit
            shipTo.validate("Address 2", format(address3.trim() + ' ' + address4.trim()).trim());
        end;

        // // City State Zip Country
        shipTo.City := copystr(city, 1, 30);
        shipTo.county := state;
        shipTo."Post Code" := zip;
        shipTo."Country/Region Code" := addressUtility.CountryToCountryCode(country);
        if shipTo."Country/Region Code" = '' then begin
            if addressUtility.ValidateUsState(shipTo.County) then begin
                shipTo."Country/Region Code" := 'US';
            end;
        end;
        shipto.Modify();
    end;


    procedure ImportQbPurchaseLines()
    var
        qbPOLine: Record "GT - QB PO Line";
        qbPOLine2: Record "GT - QB PO Line";
        qbPOHeader: Record "GT - QB PO Header";
        excelBuffer: record "Excel Buffer" temporary;
        instr: InStream;
        lastrow: Integer;
        row: Integer;
        columnNumber: Integer;
        FromFile: Text[100];
        copyValue: Text;
        d: Dialog;
        dText: label 'Create PO #1 of #2.';
        customerNo: text;
        PoNo: Text;

        listOfText: List of [Text];
        description: Text;
        descriptionTemp: Text;
        cr10: char;
        cr13: char;
        counter: integer;
        LineNo: integer;
        tb: TextBuilder;
        fm: codeunit "GT - File Management";
    begin
        cr10 := 10;
        cr13 := 13;
        if UploadIntoStream('Import QB PO Lines', '', '', FromFile, instr) then begin
            excelBuffer.OpenBookStream(instr, 'PO');
            excelBuffer.ReadSheet();
            excelBuffer.setrange("Column No.", 1);
            excelBuffer.FindLast();

            LastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);

                clear(listOfText);
                PoNo := GetText(excelBuffer, 3, row);
                if qbPOHeader.Get(PoNo) then begin
                    descriptionTemp := GetText(excelBuffer, 24, row);
                    if descriptionTemp.contains(cr10) then begin
                        listOfText := descriptionTemp.replace(cr13, '').split(cr10);
                    end
                    else
                        if descriptionTemp.contains(cr13) then begin
                            listOfText := descriptionTemp.split(cr13);
                        end else
                            listOfText.Add(descriptionTemp);

                    counter := 1;
                    foreach description in listOfText do begin
                        clear(qbPOLine);
                        qbPOLine.Init();
                        qbPOLine."PO No." := qbPOHeader."PO No.";
                        if qbPOLine2.get(qbPOHeader."PO No.", GetInteger(excelBuffer, 25, row)) then begin
                            // line already exists
                            clear(qbPOLine2);
                            qbPOLine2.setrange("PO No.", qbPOHeader."PO No.");
                            qbPOLine2.FindLast();
                            qbPOLine."Line No." := qbPOLine2."Line No." + 10000;
                        end else begin
                            qbPOLine."Line No." := GetInteger(excelBuffer, 25, row);
                        end;

                        if counter = 1 then begin
                            qbPOLine."PO Date" := GetDate(excelBuffer, 2, row);
                            qbPOLine."Amount" := GetDecimal(excelBuffer, 22, row);
                            qbPOLine."Unit Price" := GetDecimal(excelBuffer, 23, row);

                            qbPOLine."Quantity" := GetDecimal(excelBuffer, 26, row);
                            qbPOLine."Unit of Measure" := GetText(excelBuffer, 28, row);
                        end;

                        qbPOLine."Description" := description;

                        qbPOLine.insert();
                        counter += 1;
                    end;
                end
                else begin
                    tb.AppendLine('Error inserting PO line. PO not found: ' + PoNo);
                end;
            end;
            if tb.totext() <> '' then
                fm.DownloadTextAsFile(tb.ToText(), 'Estimate Lines Error Log.txt');

            d.close();

        end;
    end;


    procedure ImportQbPurchaseHeaders()
    var
        qbPOHeader: Record "GT - QB PO Header";
        qbPOHeader2: Record "GT - QB PO Header";
        excelBuffer: record "Excel Buffer" temporary;
        instr: InStream;
        lastrow: Integer;
        row: Integer;
        columnNumber: Integer;
        FromFile: Text[100];
        copyValue: Text;
        d: Dialog;
        dText: label 'Creating PO Header #1 of #2.';
        vendorNo: text;
    begin
        if UploadIntoStream('Import QB PO Headers', '', '', FromFile, instr) then begin
            excelBuffer.OpenBookStream(instr, 'PO');
            excelBuffer.ReadSheet();
            excelBuffer.setrange("Column No.", 1);
            excelBuffer.FindLast();

            LastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);

                // if GetText(excelBuffer, 46, row).trim() = 'Header' then begin
                vendorNo := GetVendorNoFromName(GetText(excelBuffer, 1, row));
                if vendorNo <> '' then begin
                    qbPOHeader2.setrange("PO No.", getText(excelBuffer, 3, row));
                    if qbPOHeader2.isEmpty then begin // prevent duplicates
                        clear(qbPOHeader);
                        qbPOHeader.Init();
                        qbPOHeader.validate("Vendor No.", vendorNo);
                        qbPOHeader.validate("Vendor Name", GetText(excelBuffer, 1, row));
                        qbPOHeader.validate("PO No.", getText(excelBuffer, 3, row));
                        qbPOHeader.validate("PO Date", getDate(excelbuffer, 2, row));

                        qbPOHeader."Bill-To address 1" := GetText(excelBuffer, 4, row);
                        qbPOHeader."Bill-To address 2" := GetText(excelBuffer, 5, row);
                        qbPOHeader."Bill-To address 3" := GetText(excelBuffer, 6, row);
                        qbPOHeader."Bill-To address 4" := GetText(excelBuffer, 7, row);
                        qbPOHeader."Bill-To city" := GetText(excelBuffer, 8, row);
                        qbPOHeader."Bill-To state" := GetText(excelBuffer, 9, row);
                        qbPOHeader."Bill-To zip" := GetText(excelBuffer, 10, row);
                        qbPOHeader."Bill-To country" := GetText(excelBuffer, 11, row);

                        qbPOHeader."ship-To address 1" := GetText(excelBuffer, 12, row);
                        qbPOHeader."ship-To address 2" := GetText(excelBuffer, 13, row);
                        qbPOHeader."ship-To address 3" := GetText(excelBuffer, 14, row);
                        qbPOHeader."ship-To address 4" := GetText(excelBuffer, 15, row);
                        qbPOHeader."ship-To city" := GetText(excelBuffer, 16, row);
                        qbPOHeader."ship-To state" := GetText(excelBuffer, 17, row);
                        qbPOHeader."ship-To zip" := GetText(excelBuffer, 18, row);
                        qbPOHeader."ship-To country" := GetText(excelBuffer, 19, row);

                        qbPOHeader."Sub Total" := GetDecimal(excelBuffer, 20, row);
                        qbPOHeader."Total Amount" := GetDecimal(excelBuffer, 22, row);

                        qbPOHeader."Ordered By" := GetText(excelBuffer, 29, row);
                        qbPOHeader.Memo := GetText(excelBuffer, 21, row);

                        qbPOHeader.insert();
                    end;
                end;
                // end;
            end;
            d.close();
        end;
    end;

    procedure ImportQbEstimateLines()
    var
        qbEstimateLine: Record "GT - QB Estimate Line";
        qbEstimateLine2: Record "GT - QB Estimate Line";
        qbEstimateHeader: Record "GT - QB Estimate Header";
        excelBuffer: record "Excel Buffer" temporary;
        instr: InStream;
        lastrow: Integer;
        row: Integer;
        columnNumber: Integer;
        FromFile: Text[100];
        copyValue: Text;
        d: Dialog;
        dText: label 'Create Estimate #1 of #2.';
        customerNo: text;
        EstimateNo: Text;
        listOfText: List of [Text];
        description: Text;
        descriptionTemp: Text;
        counter: integer;
        LineNo: integer;
        cr10: char;
        cr13: char;
        tb: TextBuilder;
        fm: codeunit "GT - File Management";
    begin
        cr10 := 10;
        cr13 := 13;
        if UploadIntoStream('Import QB Estimate Lines', '', '', FromFile, instr) then begin
            excelBuffer.OpenBookStream(instr, 'Estimates');
            excelBuffer.ReadSheet();
            excelBuffer.setrange("Column No.", 1);
            excelBuffer.FindLast();

            LastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);
                EstimateNo := GetText(excelBuffer, 2, row);

                clear(listOfText);
                if qbEstimateHeader.Get(EstimateNo) then begin
                    descriptionTemp := GetText(excelBuffer, 5, row);
                    if descriptionTemp.contains(cr10) then begin
                        listOfText := descriptionTemp.replace(cr13, '').split(cr10);
                    end
                    else
                        if descriptionTemp.contains(cr13) then begin
                            listOfText := descriptionTemp.split(cr13);
                        end else
                            listOfText.Add(descriptionTemp);

                    counter := 1;
                    foreach description in listOfText do begin
                        clear(qbEstimateLine);
                        qbEstimateLine.Init();
                        qbEstimateLine."Estimate No." := EstimateNo;

                        clear(qbEstimateLine2);
                        qbEstimateLine2.setrange("Estimate No.", EstimateNo);
                        if qbEstimateLine2.FindLast() then begin
                            qbEstimateLine."Line No." := qbEstimateLine2."Line No." + 10000;
                        end else begin
                            qbEstimateLine."Line No." := 10000;
                        end;

                        if counter = 1 then begin
                            qbEstimateLine."Amount" := GetDecimal(excelBuffer, 3, row);
                            qbEstimateLine."Quantity" := GetDecimal(excelBuffer, 6, row);

                            qbEstimateLine."Unit Price" := GetDecimal(excelBuffer, 4, row);
                            if GetText(excelBuffer, 9, row) = 'each' then
                                qbEstimateLine."Unit of Measure" := 'EA';
                        end;

                        qbEstimateLine."Description" := description;
                        if (not ((qbEstimateLine.Description = 'Sales Tax') and (qbestimateline.Amount = 0))) or
                            (not ((qbEstimateLine.Description = '') and (qbestimateline.Amount = 0) and (qbEstimateLine.quantity = 0) and (qbestimateline."unit price" = 0))) then
                            qbEstimateLine.insert();
                        counter += 1;
                    end;
                end
                else begin
                    tb.AppendLine('Error inserting invoice line. Invoice not found: ' + GetText(excelBuffer, 2, row));
                end;
            end;
            if tb.totext() <> '' then
                fm.DownloadTextAsFile(tb.ToText(), 'Estimate Lines Error Log.txt');

            d.close();
        end;
    end;

    procedure ImportQBCustomers()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        customer: Record Customer;
        series: codeunit NoSeriesManagement;
        customerName: text;
        customerNameList: List of [Text];
        tb: TextBuilder;
        phoneNumberValidator: codeunit "GT - Phone Number Validator";

        parsedPhoneNumber: text;
        PhoneExtension: Text;

        contactName: Text;
        th: codeunit "Type Helper";
        addressUtility: COdeunit "GT - Address Utility";
        util: codeunit "GT - Utilities";
        taxArea: record "Tax Area";
        salesPeople: record "Salesperson/Purchaser";
        mm: codeunit "Mail Management";
        i: Integer;
        regex: Codeunit Regex;

        d: Dialog;
        dText: label 'Creating Customer #1 of #2.';
    begin
        // if posting group doesn't exist
        if not gSetup.TestCustomerPostingGroup('USD') then begin
            Error('Customer Posting Group is not set up correctly. Please contact your system administrator.');
        end;
        if not gSetup.TestGenBusinessPostingGroup('DEFAULT') then begin
            Error('Customer Posting Group is not set up correctly. Please contact your system administrator.');
        end;

        if UploadIntoStream('Import Customer from QB', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Customers');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1); // column A is Blank
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);

            for row := 2 to lastRow do begin
                d.update();
                // if the Customer (name) contains ":" then we need to create a ship-to location for that customer.
                // we should also make sure the customer has been created.

                if GetText(excelBuffer, 7, row) = '' then begin // we are skipping children
                    clear(customer);
                    customer.Init();
                    customer.validate("No.", series.GetNextNo('CUST', Today(), true));
                    // Customer Name 4
                    customer.validate(Name, GetText(excelBuffer, 4, row));
                    customer.Insert();

                    // Main Phone 38
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    parsedPhoneNumber := GetText(excelBuffer, 38, row);
                    phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);

                    if regex.IsMatch(parsedPhoneNumber, '^\d{1}-\d{3}-\d{3}-\d{4}$', 0) then begin
                        customer.validate("Phone No.", parsedPhoneNumber);
                        CUSTOMER.validate("GT Phone Extension", PhoneExtension);
                    end
                    else
                        customer."GT QB Extra Data" += 'QB Main Phone: ' + GetText(excelBuffer, 38, row) + th.NewLine();

                    // Fax 40
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    parsedPhoneNumber := GetText(excelBuffer, 40, row);
                    phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);
                    customer.validate("Fax No.", parsedPhoneNumber);

                    // Contact// Contact name
                    clear(contactName);
                    if GetText(excelBuffer, 8, row).trim() <> '' then
                        contactName += GetText(excelBuffer, 8, row) + ' ';
                    if GetText(excelBuffer, 9, row).trim() <> '' then
                        contactName += GetText(excelBuffer, 9, row) + ' ';
                    if GetText(excelBuffer, 10, row).trim() <> '' then
                        contactName += GetText(excelBuffer, 10, row);

                    customer.validate(contact, contactName.trim()); // save it first


                    //Address
                    if GetText(excelBuffer, 12, row).trim() = customer.name.trim() then begin
                        // we dont want address1 from QB
                        customer.validate(Address, GetText(excelBuffer, 13, row).trim());
                    end
                    else begin
                        customer.validate(Address, format(GetText(excelBuffer, 12, row).trim() + ' ' + GetText(excelBuffer, 13, row).trim()).trim());
                    end;

                    if strlen(format(GetText(excelBuffer, 14, row).trim() + ' ' + GetText(excelBuffer, 15, row).trim()).trim()) > 50 then begin
                        customer.validate("GT QB Extra Data", format(GetText(excelBuffer, 14, row).trim() + ' ' + GetText(excelBuffer, 15, row).trim()).trim() + th.NewLine());
                        //copying the first 50 characters
                        customer.validate("Address 2", CopyStr(format(GetText(excelBuffer, 14, row).trim() + ' ' + GetText(excelBuffer, 15, row).trim()).trim(), 1, 50));
                    end
                    else begin
                        customer.validate("Address 2", format(GetText(excelBuffer, 14, row).trim() + ' ' + GetText(excelBuffer, 15, row).trim()).trim());
                    end;

                    // City State Zip Country
                    customer.City := copystr(GetText(excelBuffer, 16, row), 1, 30);
                    customer.county := GetText(excelBuffer, 17, row);
                    customer."Post Code" := GetText(excelBuffer, 18, row);
                    customer."Country/Region Code" := addressUtility.CountryToCountryCode(GetText(excelBuffer, 19, row));
                    if customer."Country/Region Code" = '' then begin
                        if addressUtility.ValidateUsState(customer.County) then begin
                            customer."Country/Region Code" := 'US';
                        end;
                    end;

                    customer.validate("Payment Terms Code", TermsDescToTermsCode(GetText(excelBuffer, 21, row)));

                    customer.validate("Home Page", getText(excelBuffer, 31, row));
                    //main email
                    if getText(excelBuffer, 36, row).trim() <> '' then begin
                        if mm.CheckValidEmailAddress(getText(excelBuffer, 36, row).trim()) then
                            customer.validate("E-Mail", getText(excelBuffer, 36, row))
                        else
                            customer."GT QB Extra Data" += 'QB Invalid Email: ' + GetText(excelBuffer, 36, row) + th.NewLine();
                    end;


                    // Salesperson
                    if salesPeople.get(getText(excelBuffer, 22, row)) then
                        customer.validate("Salesperson Code", getText(excelBuffer, 22, row)) // sales rep full name contains code
                    else begin
                        customer.validate("salesperson Code", 'HA');
                    end;
                    customer.validate("Credit Limit (LCY)", GetDecimal(excelBuffer, 26, row));


                    if strlen(GetText(excelBuffer, 6, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Class: ' + GetText(excelBuffer, 6, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 20, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Type: ' + GetText(excelBuffer, 20, row) + th.NewLine();

                    // taxes
                    clear(taxArea);
                    taxArea.setrange(Description, GetText(excelBuffer, 24, row));
                    if taxArea.findfirst() then begin
                        customer.validate("Tax Area Code", taxArea.code);
                    end else begin
                        customer."GT QB Extra Data" += 'QB Tax: ' + GetText(excelBuffer, 24, row) + th.NewLine();
                    end;

                    // resale # / Tax Registration #
                    if strlen(GetText(excelBuffer, 25, row).trim()) > 20 then begin
                        customer."GT QB Extra Data" += 'QB Resale #: ' + GetText(excelBuffer, 25, row) + th.NewLine();
                    end else begin
                        customer.validate("Vat Registration No.", GetText(excelBuffer, 25, row));
                    end;

                    // payment Methods
                    case GetText(excelBuffer, 27, row).trim().toupper() of
                        'CHECK':
                            customer.validate("Payment Method Code", 'CHECK');
                        'AMERICAN EXPRESS', 'DISCOVER', 'MASTERCARD', 'VISA'
                            :
                            customer.validate("Payment Method Code", 'CARD');
                        'ACH':
                            customer.validate("Payment Method Code", 'ACH');
                        'WIRE TRANSFER':
                            customer.validate("Payment Method Code", 'WIRE');
                    end;

                    // document sending profile
                    case GetText(excelBuffer, 28, row).trim().toupper() of
                        'E-MAIL', 'NONE':
                            customer.validate("Document Sending Profile", 'E-MAIL');
                        'MAIL':
                            customer.validate("Document Sending Profile", 'MAIL');
                    end;

                    customer.validate("Gen. Bus. Posting Group", 'DEFAULT');
                    customer.validate("Customer Posting Group", 'USD');


                    if strlen(GetText(excelBuffer, 29, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Work Phone: ' + GetText(excelBuffer, 29, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 41, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Alt Phone: ' + GetText(excelBuffer, 41, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 32, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Mobile: ' + GetText(excelBuffer, 32, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 39, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Alt Mobile: ' + GetText(excelBuffer, 39, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 42, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Home Phone: ' + GetText(excelBuffer, 42, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 30, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB CC Email: ' + GetText(excelBuffer, 30, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 35, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Alt Email 1: ' + GetText(excelBuffer, 35, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 33, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Alt Email 2: ' + GetText(excelBuffer, 33, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 34, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Other 1: ' + GetText(excelBuffer, 34, row) + th.NewLine();
                    if strlen(GetText(excelBuffer, 37, row).trim()) > 0 then
                        customer."GT QB Extra Data" += 'QB Other 2: ' + GetText(excelBuffer, 37, row) + th.NewLine();
                    customer.Modify();
                end;
            end;
            d.close();
        end;
    end;


    /// <summary>
    /// We used QuickBooks Advance Reporting to Create the Excel Spreadsheet
    /// </summary>    
    procedure ImportQBVendors()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        vendor: Record Vendor;
        series: codeunit NoSeriesManagement;
        vendorName: text;
        vendorNameList: List of [Text];
        tb: TextBuilder;
        phoneNumberValidator: codeunit "GT - Phone Number Validator";

        parsedPhoneNumber: text;
        PhoneExtension: Text;
        contactName: Text;
        contactLastName: Text;
        contactFirstName: Text;
        th: codeunit "Type Helper";
        addressUtility: COdeunit "GT - Address Utility";
        AddressJo: JsonObject;
        fm: Codeunit "GT - File Management";
        addressJsonText: Text;
        addressText: text;
    begin


        if UploadIntoStream('Import Vendor from QB', '', '', varFileName, inS) then begin
            excelBuffer.OpenBookStream(inS, 'Vendors');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 2); // column A is Blank
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            for row := 2 to lastRow do begin

                if GetInteger(excelBuffer, 1, row) = 1 then begin // 0 is inactive

                    clear(addressJsonText);

                    clear(vendor);
                    vendor.Init();
                    vendor.validate("No.", series.GetNextNo('VEND', Today(), true));
                    // vendor name 2
                    vendor.validate(Name, GetText(excelBuffer, 2, row));
                    vendor.Insert();

                    // Contact name
                    clear(contactName);
                    if GetText(excelBuffer, 3, row).trim() <> '' then
                        contactName += GetText(excelBuffer, 3, row) + ' ';
                    if GetText(excelBuffer, 4, row).trim() <> '' then
                        contactName += GetText(excelBuffer, 4, row) + ' ';
                    if GetText(excelBuffer, 5, row).trim() <> '' then
                        contactName += GetText(excelBuffer, 5, row);

                    vendor.validate(contact, contactName.trim()); // save it first

                    //Address
                    clear(vendorNameList);
                    if GetText(excelBuffer, 7, row).trim() = vendor.name.trim() then begin
                        // we dont want address1 from QB
                        vendor.validate(Address, GetText(excelBuffer, 8, row).trim());
                    end
                    else begin
                        vendor.validate(Address, format(GetText(excelBuffer, 7, row).trim() + ' ' + GetText(excelBuffer, 8, row).trim()).trim());
                    end;
                    vendor.validate("Address 2", format(GetText(excelBuffer, 9, row).trim() + ' ' + GetText(excelBuffer, 10, row).trim()).trim());

                    // City State Zip Country
                    vendor.City := GetText(excelBuffer, 12, row);
                    vendor.county := GetText(excelBuffer, 13, row);
                    vendor."Post Code" := GetText(excelBuffer, 14, row);
                    vendor."Country/Region Code" := addressUtility.CountryToCountryCode(GetText(excelBuffer, 15, row));
                    if vendor."Country/Region Code" = '' then begin
                        if addressUtility.ValidateUsState(vendor.County) then begin
                            vendor."Country/Region Code" := 'US';
                        end;
                    end;

                    // vendor account #
                    if strlen(GetText(excelBuffer, 17, row).trim()) > 20 then begin
                        vendor.validate("GT QB Account No.", GetText(excelBuffer, 17, row).trim());
                        vendor.validate("Our Account No.", 'See QB Account No.');
                    end
                    else begin
                        vendor.validate("Our Account No.", GetText(excelBuffer, 17, row));
                    end;

                    // Payment Terms Code
                    vendor.validate("Payment Terms Code", TermsDescToTermsCode(GetText(excelBuffer, 19, row)));

                    // Credit Limit
                    vendor.validate("GT QB Credit Limit", GetDecimal(excelBuffer, 20, row));
                    // 1099
                    vendor.validate("GT QB 1099 Eligible", GetBoolean(excelBuffer, 21, row));

                    
                    vendor.validate("Gen. Bus. Posting Group", 'DEFAULT'); // must be setup in BC
                    vendor.validate("Vendor Posting Group", 'USD'); // must be setup in BC

                    // Main Phone 12
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    parsedPhoneNumber := GetText(excelBuffer, 39, row);
                    phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);
                    vendor.validate("Phone No.", parsedPhoneNumber);
                    vendor.validate("GT Phone Extension", PhoneExtension);

                    // Main Email
                    vendor.validate("E-Mail", GetText(excelBuffer, 37, row));

                    // CC email
                    if strlen(GetText(excelBuffer, 31, row).trim()) > 0 then begin
                        vendor."GT QB Alternate Emails" := GetText(excelBuffer, 31, row) + ';';
                    end;

                    // alt emails
                    if strlen(GetText(excelBuffer, 34, row).trim()) > 0 then begin
                        vendor."GT QB Alternate Emails" += GetText(excelBuffer, 34, row) + ';';
                    end;
                    if strlen(GetText(excelBuffer, 36, row).trim()) > 0 then begin
                        vendor."GT QB Alternate Emails" += GetText(excelBuffer, 36, row) + ';';
                    end;

                    vendor."GT QB Alternate Emails" := vendor."GT QB Alternate Emails".TrimEnd(';');

                    // website
                    vendor.validate("Home Page", GetText(excelBuffer, 32, row));

                    // Fax 13
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    parsedPhoneNumber := GetText(excelBuffer, 41, row);
                    phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);
                    vendor.validate("Fax No.", parsedPhoneNumber);
                    vendor.Modify();

                    // Mobile 33
                    vendor.validate("GT QB Mobile", GetText(excelBuffer, 33, row));

                    // work phone 30
                    vendor.validate("GT QB Work Phone", GetText(excelBuffer, 30, row));
                end;
            end;
        end;
    end;



    procedure ImportQBItems()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        customer: Record Customer;
        series: codeunit NoSeriesManagement;
        customerName: text;
        customerNameList: List of [Text];
        tb: TextBuilder;
        phoneNumberValidator: codeunit "GT - Phone Number Validator";

        parsedPhoneNumber: text;
        PhoneExtension: Text;

        contactLastName: Text;
        contactFirstName: Text;
        th: codeunit "Type Helper";

        status: text;
        // itemtype: text;
        itemNumberText: text;
        itemNumber: integer;

        itemRec: record Item;
        description1: text;
        description2: text;
    begin


        if UploadIntoStream('Import Items from QB', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Items');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 2); // column A is Blank
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            for row := 2 to lastRow do begin
                status := GetText(excelBuffer, 1, row);
                // itemtype := GetText(excelBuffer, 3, row);
                if (status = '1') then begin //and (itemType = 'Inventory Part') then begin
                    itemRec.init();

                    // Item Group
                    itemRec.validate("no.", 'FG-' + GetText(excelBuffer, 3, row).PadLeft(6, '0')); // pad with 6 zeros
                    itemRec.validate("type", Enum::"Item Type"::Inventory);
                    itemRec.Insert();

                    description1 := GetText(excelBuffer, 9, row);
                    Clear(description2);
                    if StrLen(description1) > 100 then begin
                        description2 := CopyStr(description1, 101, StrLen(description1));
                        description1 := CopyStr(description1, 1, 100);
                    end;
                    itemRec.validate(Description, description1); //description
                    itemRec.validate("Description 2", description2); //description
                    itemRec.validate("Base Unit of Measure", 'EA');

                    // Inventory Group

                    //Costs and Posting
                    itemRec.validate("Costing Method", enum::"costing method"::FIFO); // Standard / Average ???
                    itemRec.validate("Gen. Prod. Posting Group", 'GL'); // General Product Posting Groups --> GL
                                                                        // todo - check if this is correct
                    if GetText(excelBuffer, 6, row) = 'Tax' then begin
                        itemRec.validate("Tax Group Code", 'TAXABLE'); // Tax Group Code --> TAXABLE / NONTAXABLE
                    end else begin
                        itemRec.validate("Tax Group Code", 'NONTAXABLE'); // Tax Group Code --> TAXABLE / NONTAXABLE
                    end;
                    itemRec.validate("Inventory Posting Group", 'FINISHED GOODS'); // General Product Posting Groups
                    itemRec.validate("Unit Cost", GetDecimal(excelBuffer, 10, row)); // Unit Cost


                    // Prices and Sales
                    itemRec.validate("Unit Price", GetDecimal(excelBuffer, 7, row)); // Unit Price
                    itemRec.Validate("Price/Profit Calculation", enum::"item price profit calculation"::"Profit=Price-Cost"); // Price/Profit Calculation
                    itemRec.validate("Sales Unit of Measure", 'EA'); // Sales Unit of Measure


                    //replenishment
                    itemRec.validate("Vendor Item No.", GetText(excelBuffer, 5, row)); // Vendor Item No.
                    itemRec.validate("Purch. Unit of Measure", 'EA'); // Purch. Unit of Measure
                    itemRec.validate("Put-away Unit of Measure Code", 'EA'); // Purch. Unit of Measure
                    itemRec.modify(true);
                end;

            end;
        end;
    end;


    internal procedure ImportQBCustomerContacts()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        contact: Record Contact;
        companyContact: record Contact;
        series: codeunit NoSeriesManagement;
        tb: TextBuilder;
        phoneNumberValidator: codeunit "GT - Phone Number Validator";

        parsedPhoneNumber: text;
        PhoneExtension: Text;
        contactName: Text;
        contactLastName: Text;
        contactFirstName: Text;
        th: codeunit "Type Helper";
        addressUtility: COdeunit "GT - Address Utility";
        AddressJo: JsonObject;
        fm: Codeunit "GT - File Management";
        addressJsonText: Text;
        addressText: text;
        companyName: Text;
        mm: codeunit "Mail Management";
        updatingContact: Boolean;
    begin
        if UploadIntoStream('Import QB Cust & Vend Contacts', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Contacts');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            for row := 2 to lastRow do begin
                companyName := GetText(excelBuffer, 1, row).trim();
                companyContact.setrange(name, companyName);
                companyContact.setrange(Type, enum::"Contact Type"::Company);
                clear(updatingContact);

                if companyContact.findfirst() then begin
                    // we found the company
                    contact.setrange("Company No.", companyContact."No.");
                    contact.setrange("First Name", GetText(excelBuffer, 3, row));
                    contact.setrange(Surname, GetText(excelBuffer, 5, row));
                    contact.setrange(Type, enum::"Contact Type"::Person);
                    if contact.FindFirst() then begin
                        // we found the contact -- we will update lower down
                        tb.Appendline();
                        tb.AppendLine('Contact already exists for ' + format(contact."Contact Business Relation") + ': ' + companyName + ': ' + contact.Name);
                        updatingContact := true;
                    end
                    else begin
                        //create new  contact
                        clear(contact);
                        contact.Init();
                        contact.Validate("No.", series.GetNextNo('CONT', Today, true));
                        contact.validate(Type, enum::"Contact Type"::Person);
                        contact.validate("Company No.", companyContact."No.");
                        contact.validate(name, GetText(excelBuffer, 3, row).trim() + ' ' + GetText(excelBuffer, 5, row).trim());
                        contact.Insert();
                    end;

                    // Middle Name
                    if GetText(excelBuffer, 4, row).trim() <> '' then begin
                        if updatingcontact then begin
                            tb.appendline('- Middle :' + contact."Middle Name" + ' --> ' + GetText(excelBuffer, 4, row).Trim());
                        end;
                        contact.validate("Middle Name", GetText(excelBuffer, 4, row).trim());
                    end;
                    // Job Title
                    if GetText(excelBuffer, 6, row).trim() <> '' then begin
                        if updatingcontact then begin
                            tb.appendline('- Job Title :' + contact."Job Title" + ' --> ' + GetText(excelBuffer, 6, row).Trim());
                        end;
                        if strlen(GetText(excelBuffer, 6, row).trim()) > 30 then begin
                            contact.validate("Job Title", copystr(GetText(excelBuffer, 6, row).trim(), 1, 30));
                            contact."GT QB Extra Data" += 'Job Tile: ' + GetText(excelBuffer, 6, row).trim() + th.NewLine();
                        end
                        else
                            contact.validate("Job Title", GetText(excelBuffer, 6, row).trim());
                    end;
                    // Email
                    if GetText(excelBuffer, 7, row).trim() <> '' then begin
                        if updatingcontact then begin
                            tb.appendline('- Email :' + contact."E-mail" + ' --> ' + GetText(excelBuffer, 7, row).Trim());
                        end;

                        if mm.CheckValidEmailAddress(getText(excelBuffer, 7, row).trim()) then
                            contact.validate("E-Mail", getText(excelBuffer, 7, row))
                        else
                            contact."GT QB Extra Data" += 'QB Invalid Email: ' + GetText(excelBuffer, 7, row) + th.NewLine();
                    end;

                    // Other email
                    if GetText(excelBuffer, 8, row).trim() <> '' then begin
                        contact."GT QB Extra Data" += 'Email 2: ' + GetText(excelBuffer, 8, row).trim() + th.NewLine();
                    end;

                    // Web Site
                    if GetText(excelBuffer, 9, row).trim() <> '' then begin
                        if updatingcontact then begin
                            tb.appendline('- Home Page :' + contact."Home Page" + ' --> ' + GetText(excelBuffer, 9, row).Trim());
                        end;
                        contact.validate("Home Page", GetText(excelBuffer, 9, row).trim());
                    end;

                    // Work Phone
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    if GetText(excelBuffer, 10, row).trim() <> '' then begin
                        parsedPhoneNumber := GetText(excelBuffer, 10, row).trim();
                        phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);
                        if updatingcontact then begin
                            tb.appendline('- Phone No.:' + contact."Phone No." + 'Ex: ' + contact."Extension No." + ' --> ' + parsedPhoneNumber + ' Ex:' + PhoneExtension);
                        end;
                        contact.validate("Phone No.", parsedPhoneNumber);
                        if PhoneExtension <> '' then
                            contact.validate("Extension No.", PhoneExtension);
                    end;
                    // Mobile
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    if GetText(excelBuffer, 11, row).trim() <> '' then begin
                        parsedPhoneNumber := GetText(excelBuffer, 11, row).trim();
                        phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);
                        if updatingcontact then begin
                            tb.appendline('- Mobile Phone No.:' + contact."Mobile Phone No." + ' --> ' + parsedPhoneNumber);
                        end;
                        contact.validate("Mobile Phone No.", parsedPhoneNumber);
                    end;

                    // Work Fax
                    clear(parsedPhoneNumber);
                    clear(PhoneExtension);
                    if GetText(excelBuffer, 12, row).trim() <> '' then begin
                        parsedPhoneNumber := GetText(excelBuffer, 12, row).trim();
                        phoneNumberValidator.CleanQuickBooksPhoneNumber(parsedPhoneNumber, PhoneExtension);
                        if updatingcontact then begin
                            tb.appendline('- Fax No.:' + contact."Fax No." + ' --> ' + parsedPhoneNumber);
                        end;
                        contact.validate("Mobile Phone No.", parsedPhoneNumber);
                    end;
                    contact.Modify();
                end;
            end;
            fm.DownloadTextAsFile(tb.ToText(), 'Customer Contacts Modifications.txt');
        end;
    end;

    internal procedure ImportQBCustomerShipToAddresses()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        customer: Record Customer;
        series: codeunit NoSeriesManagement;
        tb: TextBuilder;
        phoneNumberValidator: codeunit "GT - Phone Number Validator";
        custNameList: List of [Text];

        parsedPhoneNumber: text;
        PhoneExtension: Text;
        contactName: Text;
        contactLastName: Text;
        contactFirstName: Text;
        th: codeunit "Type Helper";
        addressUtility: COdeunit "GT - Address Utility";
        AddressJo: JsonObject;
        fm: Codeunit "GT - File Management";
        addressJsonText: Text;
        addressText: text;
        companyName: Text;
        mm: codeunit "Mail Management";
        updatingContact: Boolean;



        shipToName: Text;
        address1: Text;
        address2: Text;
        address3: Text;
        address4: Text;
        city: Text;
        state: Text;
        zip: Text;
        country: Text;
    begin
        if UploadIntoStream('Import QB Cust Ship-Tos', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Customer Ship-Tos');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            for row := 2 to lastRow do begin
                // check if ship to is blank
                if GetText(excelBuffer, 3, row).Trim() <> '' then begin
                    // find customer by name
                    // handle colon names
                    if GetText(excelBuffer, 2, row).Contains(':') then begin
                        custNameList := GetText(excelBuffer, 2, row).trim().Split(':');
                        customer.SetRange(Name, custNameList.get(1).trim());
                    end
                    else begin
                        customer.SetRange(Name, GetText(excelBuffer, 2, row).Trim());
                    end;

                    if customer.findfirst() then begin
                        // create ship to
                        shipToName := GetText(excelBuffer, 3, row).Trim();
                        address1 := GetText(excelBuffer, 4, row).Trim();
                        address2 := GetText(excelBuffer, 5, row).Trim();
                        address3 := GetText(excelBuffer, 6, row).Trim();
                        address4 := GetText(excelBuffer, 7, row).Trim();
                        city := GetText(excelBuffer, 9, row).Trim();
                        state := GetText(excelBuffer, 10, row).Trim();
                        zip := GetText(excelBuffer, 11, row).Trim();
                        country := GetText(excelBuffer, 12, row).Trim();

                        CreateCustomerShipTo(customer."No.", shiptoname, address1, address2, address3, address4, city, state, zip, country);
                    end;
                end;
            end;
        end;
    end;

    internal procedure ImportQBInvoiceLines()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        customer: Record Customer;
        qbInvoice: Record "GT - QB Invoice Header";
        qbLine: Record "GT - QB Invoice Line";
        qbLine2: Record "GT - QB Invoice Line";
        tb: TextBuilder;
        th: codeunit "Type Helper";
        fm: Codeunit "GT - File Management";
        InvoiceNo: text;
        d: Dialog;
        dText: label 'Creating Invoice Line #1 of #2.';
        listOfText: List of [Text];
        description: Text;
        descriptionTemp: Text;
        counter: integer;
        LineNo: integer;
        cr10: char;
        cr13: char;

    begin
        cr10 := 10;
        cr13 := 13;
        if UploadIntoStream('Import QB Invoice Lines', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Invoice Lines');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                if not (GetText(excelBuffer, 7, row) in ['Cost of Goods Sold', 'Other Current Asset']) then begin

                    d.Update(1, row);
                    InvoiceNo := GetText(excelBuffer, 2, row);

                    if qbinvoice.Get(InvoiceNo) then begin
                        descriptionTemp := GetText(excelBuffer, 5, row);
                        //check for split
                        //if only one then insert into list
                        clear(listOfText);
                        descriptionTemp := descriptionTemp.replace('_x000D_', '');
                        if descriptionTemp.contains(cr10) then begin
                            listOfText := descriptionTemp.replace(cr13, '').split(cr10);
                        end
                        else
                            if descriptionTemp.contains(cr13) then begin
                                listOfText := descriptionTemp.split(cr13);
                            end else
                                listOfText.Add(descriptionTemp);

                        counter := 1;
                        foreach description in listOfText do begin
                            clear(qbLine);
                            qbLine.Init();
                            qbLine."Invoice No." := qbInvoice."Invoice No.";

                            clear(qbLine2);
                            qbLine2.setrange("Invoice No.", qbInvoice."Invoice No.");
                            if qbLine2.FindLast() then begin
                                qbLine."Line No." := qbLine2."Line No." + 10000;
                            end
                            else begin
                                qbLine."Line No." := 10000;
                            end;


                            if counter = 1 then begin
                                qbline.Amount := GetDecimal(excelBuffer, 3, row);
                                qbLine."Quantity" := GetDecimal(excelBuffer, 6, row);
                                qbLine."Unit Price" := GetDecimal(excelBuffer, 4, row);
                                qbLine."Account Type" := GetText(excelBuffer, 7, row);
                                qbLine."Account Full Name" := GetText(excelBuffer, 8, row);
                            end;

                            qbLine."Description" := description;

                            if (qbLine."Description" = 'Non-Taxable') and (qbline.amount = 0) then
                                qbLine."Description" := ''//this is irrelevant
                            else begin
                                if not qbline.insert() then
                                    tb.AppendLine('Error inserting invoice line: ' + qbInvoice."Invoice No." + ' - ' + format(qbLine."Line No."));
                            end;
                            counter += 1;
                        end;
                    end
                    else begin
                        tb.AppendLine('Error inserting invoice line. Invoice not found: ' + GetText(excelBuffer, 2, row));
                    end;
                end;
            end;

            if tb.totext() <> '' then
                fm.DownloadTextAsFile(tb.ToText(), 'Invoice Lines Error Log.txt');

            d.Close();
        end;
    end;

    procedure CreateChartOfAccounts()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        glAccount: Record "G/L Account";

        tb: TextBuilder;
        th: codeunit "Type Helper";
        fm: Codeunit "GT - File Management";

        d: Dialog;
        dText: label 'Creating GL Accounts Line #1 of #2.';
        GlAccountCategory: record "G/L Account Category";
    begin
        if UploadIntoStream('Import Globe GL Accounts', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Sheet1');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);

                glAccount.Init();
                glAccount."No." := GetText(excelBuffer, 1, row);
                glAccount."Name" := GetText(excelBuffer, 2, row);

                // Income/Balance
                if GetText(excelBuffer, 5, row) = 'Balance Sheet' then
                    glaccount."Income/Balance" := glAccount."Income/Balance"::"Balance Sheet"
                else
                    glaccount."Income/Balance" := glAccount."Income/Balance"::"Income Statement";
                // Account Category
                case GetText(excelBuffer, 6, row) of
                    'Assets':
                        glAccount."Account Category" := glAccount."Account Category"::"Assets";
                    'Liabilities':
                        glAccount."Account Category" := glAccount."Account Category"::Liabilities;
                    'Equity':
                        glAccount."Account Category" := glAccount."Account Category"::Equity;
                    'Income':
                        glAccount."Account Category" := glAccount."Account Category"::Income;
                    'Expense':
                        glAccount."Account Category" := glAccount."Account Category"::Expense;
                    'Cost of Goods Sold':
                        glAccount."Account Category" := glAccount."Account Category"::"Cost of Goods Sold";
                end;
                //Account Sub-Category
                GlAccountCategory.SetRange(Description, GetText(excelBuffer, 7, row));
                if GlAccountCategory.findfirst() then
                    glAccount."Account Subcategory Entry No." := GlAccountCategory."Entry No.";
                // Account Type
                glAccount."Account Type" := Enum::"G/L Account Type"::Posting;
                glAccount.Insert();
            end;
        end;
    end;

    procedure ImportQBInvoiceHeaders()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        customer: Record Customer;
        qbInvoice: Record "GT - QB Invoice Header";
        tb: TextBuilder;
        th: codeunit "Type Helper";
        fm: Codeunit "GT - File Management";
        customerNo: text;
        d: Dialog;
        dText: label 'Creating Invoice #1 of #2.';
        salesOrderNo: Text;
        salesOrderNoInt: integer;
    begin
        if UploadIntoStream('Import QB Invoice Header', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Invoices');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);

                if GetText(excelBuffer, 38, row).trim() = 'Header' then begin
                    customerNo := GetCustomerNoFromName(GetText(excelBuffer, 63, row));
                    if customerNo <> '' then begin
                        qbInvoice.Init();
                        qbInvoice."Invoice No." := GetText(excelBuffer, 2, row);
                        qbInvoice."Invoice Date" := GetDate(excelBuffer, 1, row);
                        qbInvoice."Customer No." := customerNo;
                        qbInvoice."Customer Name Full" := GetText(excelBuffer, 63, row);

                        salesOrderNo := GetText(excelBuffer, 35, row);
                        salesOrderNo := salesOrderNo.replace('P', ''); // we found 1 invoice where the sales order started with P
                        if salesOrderNo <> '' then
                            evaluate(qbInvoice."Sales Order No.", salesOrderNo);

                        // qbinvoice."Sales Order No." := GetInteger(excelBuffer, 35, row);
                        // qbInvoice."Customer Message" := GetText(excelBuffer, 273, row);
                        qbinvoice."Invoice Memo" := GetText(excelBuffer, 30, row);
                        // qbInvoice.Tax := GetText(excelBuffer, 37, row);
                        qbInvoice."Customer Tax Code" := GetText(excelBuffer, 70, row);
                        qbInvoice."Sub Total" := GetDecimal(excelBuffer, 23, row);
                        qbInvoice."Ship Date" := GetDate(excelBuffer, 22, row);
                        qbInvoice."PO Number" := GetText(excelBuffer, 19, row);

                        qbInvoice."Bill-To address 1" := GetText(excelBuffer, 3, row);
                        qbInvoice."Bill-To address 2" := GetText(excelBuffer, 4, row);
                        qbInvoice."Bill-To address 3" := GetText(excelBuffer, 5, row);
                        qbInvoice."Bill-To address 4" := GetText(excelBuffer, 6, row);
                        qbInvoice."Bill-To city" := GetText(excelBuffer, 7, row);
                        qbInvoice."Bill-To state" := GetText(excelBuffer, 8, row);
                        qbInvoice."Bill-To zip" := GetText(excelBuffer, 9, row);
                        qbInvoice."Bill-To country" := GetText(excelBuffer, 10, row);

                        qbInvoice."Ship-To address 1" := GetText(excelBuffer, 11, row);
                        qbInvoice."Ship-To address 2" := GetText(excelBuffer, 12, row);
                        qbInvoice."Ship-To address 3" := GetText(excelBuffer, 13, row);
                        qbInvoice."Ship-To address 4" := GetText(excelBuffer, 14, row);
                        qbInvoice."Ship-To city" := GetText(excelBuffer, 15, row);
                        qbInvoice."Ship-To state" := GetText(excelBuffer, 16, row);
                        qbInvoice."Ship-To zip" := GetText(excelBuffer, 17, row);
                        qbInvoice."Ship-To country" := GetText(excelBuffer, 18, row);

                        qbInvoice."Total Amount" := GetDecimal(excelBuffer, 36, row);
                        qbInvoice.industry := GetText(excelBuffer, 64, row);
                        qbInvoice."Terms Description" := GetText(excelBuffer, 65, row);
                        qbInvoice."Shipment Method" := GetText(excelBuffer, 66, row);
                        qbInvoice."Sales Rep Name" := GetText(excelBuffer, 67, row);

                        if not qbInvoice.insert() then
                            tb.AppendLine('Error inserting invoice: ' + qbInvoice."Invoice No.");
                    end
                    else begin
                        tb.AppendLine('Error inserting invoice. Customer not found: ' + GetText(excelBuffer, 63, row));
                    end;
                end;
            end;
            fm.DownloadTextAsFile(tb.ToText(), 'Error Log.txt');

            d.Close();
        end;
    end;

    procedure ImportQBEstimateHeaders()
    var
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: text;
        row: Integer;
        lastrow: Integer;
        inS: InStream;
        customer: Record Customer;
        qbEstimate: Record "GT - QB Estimate Header";
        tb: TextBuilder;
        th: codeunit "Type Helper";
        fm: Codeunit "GT - File Management";
        customerNo: text;
        d: Dialog;
        dText: label 'Creating Invoice #1 of #2.';
    begin
        if UploadIntoStream('Import QB Estimate Headers', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'Estimates');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);

                customerNo := GetCustomerNoFromName(GetText(excelBuffer, 63, row));
                if customerNo <> '' then begin
                    qbEstimate.Init();
                    qbEstimate."Estimate No." := GetText(excelBuffer, 2, row);
                    qbEstimate."Estimate Date" := GetDate(excelBuffer, 1, row);
                    qbEstimate."Customer No." := customerNo;
                    qbEstimate."Customer Name Full" := GetText(excelBuffer, 63, row);
                    // qbEstimate."Sales Order No." := GetInteger(excelBuffer, 35, row);
                    // qbInvoice."Customer Message" := GetText(excelBuffer, 273, row);
                    qbEstimate.Memo := GetText(excelBuffer, 30, row);
                    // qbInvoice.Tax := GetText(excelBuffer, 37, row);
                    // qbEstimate.tax := GetText(excelBuffer, 70, row);
                    qbEstimate."Sub Total" := GetDecimal(excelBuffer, 23, row);
                    // qbEstimate."Ship Date" := GetDate(excelBuffer, 22, row);
                    // qbEstimate."PO Number" := GetText(excelBuffer, 19, row);

                    qbEstimate."Bill-To address 1" := GetText(excelBuffer, 3, row);
                    qbEstimate."Bill-To address 2" := GetText(excelBuffer, 4, row);
                    qbEstimate."Bill-To address 3" := GetText(excelBuffer, 5, row);
                    qbEstimate."Bill-To address 4" := GetText(excelBuffer, 6, row);
                    qbEstimate."Bill-To city" := GetText(excelBuffer, 7, row);
                    qbEstimate."Bill-To state" := GetText(excelBuffer, 8, row);
                    qbEstimate."Bill-To zip" := GetText(excelBuffer, 9, row);
                    qbEstimate."Bill-To country" := GetText(excelBuffer, 10, row);

                    qbEstimate."Ship-To address 1" := GetText(excelBuffer, 11, row);
                    qbEstimate."Ship-To address 2" := GetText(excelBuffer, 12, row);
                    qbEstimate."Ship-To address 3" := GetText(excelBuffer, 13, row);
                    qbEstimate."Ship-To address 4" := GetText(excelBuffer, 14, row);
                    qbEstimate."Ship-To city" := GetText(excelBuffer, 15, row);
                    qbEstimate."Ship-To state" := GetText(excelBuffer, 16, row);
                    qbEstimate."Ship-To zip" := GetText(excelBuffer, 17, row);
                    qbEstimate."Ship-To country" := GetText(excelBuffer, 18, row);

                    qbEstimate."Total Amount" := GetDecimal(excelBuffer, 36, row);
                    qbEstimate.industry := GetText(excelBuffer, 64, row);
                    qbEstimate."Sales Rep Name" := GetText(excelBuffer, 67, row);

                    if not qbEstimate.insert() then
                        tb.AppendLine('Error inserting invoice: ' + qbEstimate."Estimate No.");
                end
                else begin
                    tb.AppendLine('Error inserting invoice. Customer not found: ' + GetText(excelBuffer, 63, row));
                end;
            end;
            fm.DownloadTextAsFile(tb.ToText(), 'Error Log.txt');

            d.Close();
        end;
    end;

    procedure GetCustomerNoFromName(customerName: Text) result: code[20]
    var
        customer: record customer;
        custNameList: List of [Text];
    begin
        if customerName.Contains(':') then begin
            custNameList := customerName.trim().Split(':');
            customer.SetRange(Name, custNameList.get(1).trim());
        end
        else begin
            customer.SetRange(Name, customerName.Trim());
        end;

        if customer.findfirst() then begin
            result := customer."No.";
        end;
    end;

    procedure GetVendorNoFromName(vendorName: Text) result: code[20]
    var
        vendor: record vendor;
        vendNameList: List of [Text];
    begin
        if vendorName.Contains(':') then begin
            vendNameList := vendorName.trim().Split(':');
            vendor.SetRange(Name, vendNameList.get(1).trim());
        end
        else begin
            vendor.SetRange(Name, vendorName.Trim());
        end;

        if vendor.findfirst() then begin
            result := vendor."No.";
        end;
    end;

    procedure importPostCodes()
    var
        excelBuffer: record "Excel Buffer" temporary;
        lastRow: Integer;
        d: Dialog;
        dText: label 'Importing Post Codes #1 of #2';
        varFileName: Text;
        Ins: InStream;
        row: Integer;

        postCode: record "Post Code";
        addressUtil: codeunit "GT - Address Utility";
    begin
        if UploadIntoStream('Import QB PV Post Codes', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'City-State-Zip');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            d.Open(dText, row, lastRow);
            for row := 2 to lastRow do begin
                d.Update(1, row);
                clear(postCode);
                postCode.setrange("city", GetText(excelBuffer, 1, row).trim());
                postCode.setrange(Code, GetText(excelBuffer, 3, row).trim());

                if postCode.isEmpty() then begin
                    addressUtil.PopulatePostCodeTable(GetText(excelBuffer, 3, row).trim(),
                                                       GetText(excelBuffer, 1, row).trim(),
                                                       GetText(excelBuffer, 2, row).trim(),
                                                       GetText(excelBuffer, 4, row).trim());
                end;
            end;
            d.close();
        end;
    end;

    procedure CreateItemCategories()
    var
        itemCategory: record "Item Category";
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: Text;
        lastRow: Integer;
        ins: InStream;
        row: Integer;
    begin
        if UploadIntoStream('Import Item Categories', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'BC');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            for row := 2 to lastRow do begin
                clear(itemCategory);
                itemCategory.init();
                itemCategory.Code := GetText(excelBuffer, 1, row).Trim();
                itemCategory."Parent Category" := GetText(excelBuffer, 2, row).Trim();
                itemCategory.Description := GetText(excelBuffer, 3, row).Trim();
                itemCategory.Indentation := GetDecimal(excelBuffer, 4, row);
                itemCategory.Insert();
            end;
        end;
    end;

    procedure CreateNonStockItems()
    var
        Item: record "Item";
        excelBuffer: record "Excel Buffer" temporary;
        varFileName: Text;
        lastRow: Integer;
        ins: InStream;
        row: Integer;
    begin
        if UploadIntoStream('Import NS-Items', '', '', varFileName, inS) then begin

            excelBuffer.OpenBookStream(inS, 'BC');
            excelBuffer.ReadSheet();
            excelBuffer.SetRange("Column No.", 1);
            excelBuffer.FindLast();

            lastRow := excelBuffer."Row No.";
            excelBuffer.Reset();

            for row := 2 to lastRow do begin
                if GetText(excelBuffer, 6, row).Trim() <> '' then begin

                    clear(Item);
                    Item.init();
                    Item."No." := GetText(excelBuffer, 6, row).Trim();
                    Item.Description := GetText(excelBuffer, 3, row).Trim();
                    Item."Item Category Code" := GetText(excelBuffer, 1, row).Trim();
                    item.Type := Enum::"item type"::"Non-Inventory";
                    Item.Insert();
                    item.validate("Base Unit of Measure", 'EA');
                    item.validate("Sales Unit of Measure", 'EA');
                    item.validate("Purch. Unit of Measure", 'EA');
                    item.validate("Put-away Unit of Measure Code", 'EA');
                    item."Gen. Prod. Posting Group" := 'GL';
                    item."Tax Group Code" := 'TAXABLE';
                    Item.modify();
                end;
            end;
        end;
    end;
}
