pageextension 50100 "PurchaseOrderExt" extends "Purchase Order" //50
{
    actions
    {
        addafter(IncomingDocument)
        {
            action(ImportSeriesNo)
            {
                Caption = 'Import Series No.', Comment = 'FRA="Import N° Série"';
                ApplicationArea = all;
                Image = ImportExcel;

                trigger OnAction()
                var
                    RecLPurchaseLine: Record "Purchase Line";
                begin
                    RecLPurchaseLine.SetRange("Document No.", "No.");
                    IF RecLPurchaseLine.IsEmpty then
                        Error(CstG0001, "No.");

                    ReadExcelSheet;
                    ImportExcelData;
                end;
            }
        }
    }

    var
        CstG0001: Label 'There are no lines in the command %1', Comment = 'FRA="Il n’y a pas de lignes dans la commande %1"';
        CstG0002: Label 'Order N° %1 different from current order n° %2', Comment = 'FRA="N° de Commande %1 différent de la N° de commande actuelle %2"';
        CstG0003: Label 'Item %1 does not exist in the order', Comment = 'FRA="L''article %1 n''existe pas dans la commande"';
        TxtGFileName: Text[100];
        TxtGSheetName: Text[100];
        TempExcelBuffer: Record "Excel Buffer" temporary;
        Rows: Integer;
        EntryNo: Integer;
        RecGTrackingSpecification: record "Tracking Specification";
        RecGReservationEntry: Record "Reservation Entry";

    procedure ReadExcelSheet()
    var
        CduLFileMgt: codeunit "File Management";
        IStream: InStream;
        FromFile: Text[100];
    begin
        UploadIntoStream('', '', '', FromFile, IStream);
        IF FromFile <> '' then begin
            TxtGFileName := CduLFileMgt.GetFileName(FromFile);
            TxtGSheetName := TempExcelBuffer.SelectSheetsNameStream(IStream);
        end else
            Error('No File');

        TempExcelBuffer.Reset();
        TempExcelBuffer.DeleteAll();
        TempExcelBuffer.OpenBookStream(IStream, TxtGSheetName);
        TempExcelBuffer.ReadSheet();
    end;

    procedure ImportExcelData()
    var
        RowNo: Integer;
        RecLItem: Record item;
        RecLPurchaseLine: Record "Purchase Line";
        CodLOrderNo: Code[20];
        CodLItemNo: Code[20];
        CodLSerialNo: Code[50];
        RecLReservationEntry: Record "Reservation Entry";
    begin
        TempExcelBuffer.Reset();
        TempExcelBuffer.SetRange("Column No.", 1);
        If TempExcelBuffer.FindFirst() then
            repeat
                Rows := Rows + 1;
            until TempExcelBuffer.Next() = 0;

        For RowNo := 2 to Rows do begin
            //Vérifier même N° Commande
            Evaluate(CodLOrderNo, GetValueAtIndex(RowNo, 1));
            IF "No." <> CodLOrderNo then
                Error(CstG0002, CodLOrderNo, "No.")
            else begin
                //Vérifier N° Article existe dans les lignes Commandes
                Evaluate(CodLItemNo, GetValueAtIndex(RowNo, 2));
                Evaluate(CodLSerialNo, GetValueAtIndex(RowNo, 3));
                RecLPurchaseLine.SetRange("Document No.", "No.");
                RecLPurchaseLine.SetRange(Type, RecLPurchaseLine.Type::Item);
                RecLPurchaseLine.SetRange("No.", CodLItemNo);
                IF Not RecLPurchaseLine.IsEmpty then begin
                    RecLPurchaseLine.FindFirst();
                    IF RecLItem.get(CodLItemNo) then
                        //Vérifier que l'article a un code traçabilié
                        IF RecLItem."Item Tracking Code" <> '' then begin
                            IF RecLReservationEntry.FindLast() then
                                EntryNo := RecLReservationEntry."Entry No." + 1
                            else
                                EntryNo := 1;

                            RecGReservationEntry.Init;
                            RecGReservationEntry."Entry No." := EntryNo;
                            RecGReservationEntry.validate("Item No.", CodLItemNo);
                            RecGReservationEntry.validate("Quantity (Base)", 1);
                            RecGReservationEntry."Reservation Status" := RecGReservationEntry."Reservation Status"::Surplus;
                            RecGReservationEntry."Source Type" := Database::"Purchase Line";
                            RecGReservationEntry."Source Subtype" := RecGReservationEntry."Source Subtype"::"1";
                            RecGReservationEntry."Source ID" := CodLOrderNo;
                            RecGReservationEntry."Source Prod. Order Line" := 0;
                            RecGReservationEntry."Source Ref. No." := RecLPurchaseLine."Line No.";
                            RecGReservationEntry."Item Tracking" := RecGReservationEntry."Item Tracking"::"Serial No.";
                            RecGReservationEntry."Serial No." := CodLSerialNo;
                            RecGReservationEntry.Positive := true;
                            RecGReservationEntry.Insert();
                        end;
                end else
                    Error('CstG0003', CodLItemNo);
            end;
        end;
    end;

    local procedure GetValueAtIndex(RowNo: Integer; ColNo: Integer): Text
    begin
        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text");

    end;
}