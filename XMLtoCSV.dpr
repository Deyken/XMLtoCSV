program XMLtoCSV;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils, Winapi.Windows, Winapi.Messages, System.Variants, System.Classes, System.IOUtils, Xml.xmldom,
  Xml.XMLIntf, Xml.XMLDoc, System.UITypes, ActiveX, Math;

//___________________________________________________________________________________________________________________________________________

function AccessFile(file_name : String) : Boolean;
begin
  Result := FileExists(file_name);
end;

//___________________________________________________________________________________________________________________________________________

procedure ParseXMLData(in_file : String; var out_file : String; AOwner : TComponent; Delim : String = '|');
var
  T1, T2, T3, T4, T5 : Integer;
  i, j, k, l, m, aa : Integer;
  DataLine : TSTringList;
  DataString : String;
  DataOutput : TSTringList;
  XML1 : TXMLDocument;
  N1, N2, N3, N4, N5 : IXMLNode;
  OutLOG : TStringlist;
  XMLFileName, CSVFilename, LOGFilename : String;
  Hr, Mn, Sec, mSec : Word;
  StartTime, EndTime : TTime;

const
  OutDIR : String = 'C:\CSVOutput';

begin
  try
    StartTime := Now;

    //Create Stringlists
    DataLine := TSTringlist.Create;
    DataOutput := TSTringlist.Create;
    OutLOG := TStringList.Create;
    OutLOG.Duplicates := dupIgnore;
    DataString := '';

    try
      //Step 1: LOAD THE FILE INTO MEMORY
      if FileExists(in_file) then
        begin
          CoInitialize(nil);
          Assert(AOwner <> nil, 'ERROR: Issue with creating the OWNER of the XML Document!');
          XML1 := TXMLDocument.Create(AOwner);
          XML1.LoadFromFile(in_file);
          OutLOG.Add('New XML Document loaded from file: ' + in_file);
          OutLOG.Add('');
          OutLOG.Add('XML Data parsing commenced at ' + DateTimeToStr(Now));
          OutLOG.Add('');
          OutLOG.Add('DATA STRUCTURE:');
          OutLOG.Add('');

          //START LOOPING THROUGH THIS DOCUMENT FOR PARSING
          for T1 := 0 to XML1.ChildNodes.Count - 1 do
            begin
              N1 := XML1.ChildNodes[T1];
              OutLOG.Add('Tier 1: ' + N1.NodeName + ' with ' + IntToStr(N1.ChildNodes.Count) + ' Child Nodes and ' + IntToStr(N1.AttributeNodes.Count) + ' Fields');

              for T2 := 0 to XML1.ChildNodes[T1].ChildNodes.Count - 1 do
                begin
                  N2 := XML1.ChildNodes[T1].ChildNodes[T2];
                  OutLOG.Add(#9 + 'Tier 2: ' + N2.NodeName + ' with ' + IntToStr(N2.ChildNodes.Count) + ' Child Nodes and ' + IntToStr(N2.AttributeNodes.Count) + ' Fields');

                  //Parse all Attibutes
                  for i := 0 to N2.AttributeNodes.Count - 1 do
                    begin
                      //START BULDING THE DATA LINE HERE
                      OutLOG.Add(#9 + '  Field: ' + N2.AttributeNodes[i].NodeName);
                      DataLine.Add(N2.AttributeNodes[i].Text);
                    end;

                  for T3 := 0 to XML1.ChildNodes[T1].ChildNodes[T2].ChildNodes.Count - 1 do
                    begin
                      N3 :=  XML1.ChildNodes[T1].ChildNodes[T2].ChildNodes[T3];
                      OutLOG.Add(#9+#9 + 'Tier 3: ' + N3.NodeName + ' with ' + IntToStr(N3.ChildNodes.Count) + ' Child Nodes and ' + IntToStr(N3.AttributeNodes.Count) + ' Fields');

                      for j := 0 to N3.AttributeNodes.Count - 1 do
                        begin
                          OutLOG.Add(#9+#9 + '  Field: ' + N3.AttributeNodes[j].NodeName);
                          DataLine.Add(N3.AttributeNodes[j].Text);
                        end;

                      for T4 := 0 to XML1.ChildNodes[T1].ChildNodes[T2].ChildNodes[T3].ChildNodes.Count - 1 do
                        begin
                          N4 := XML1.ChildNodes[T1].ChildNodes[T2].ChildNodes[T3].ChildNodes[T4];
                          OutLOG.Add(#9+#9+#9 + 'Tier 4: ' + N4.NodeName + ' with ' + IntToStr(N4.ChildNodes.Count) + ' Child Nodes and ' + IntToStr(N4.AttributeNodes.Count) + ' Fields');

                          for k := 0 to N4.AttributeNodes.Count - 1 do
                            begin
                              OutLOG.Add(#9+#9+#9 + '  Field: ' + N4.AttributeNodes[k].NodeName);
                              DataLine.Add(N4.AttributeNodes[k].Text);
                            end;

                          //TEST FOR AND ACCESS A FIFTH TIER, SHOULD THE LATTER BE PRESENT IN THE XML DOCUMENT
                          if N4.ChildNodes.Count > 0 then begin
                            for T5 := 0 to N4.ChildNodes.Count - 1 do
                              begin
                                N5 := XML1.ChildNodes[T1].ChildNodes[T2].ChildNodes[T3].ChildNodes[T4].ChildNodes[T5];
                                OutLOG.Add(#9+#9+#9+#9 + 'Tier 5: ' + N5.NodeName + ' with ' + IntToStr(N5.ChildNodes.Count) + ' Child Nodes and ' + IntToStr(N5.AttributeNodes.Count) + ' Fields');

                                for m := 0 to N5.AttributeNodes.Count - 1 do
                                  begin
                                    OutLOG.Add(#9+#9+#9+#9 + '  Field: ' + N5.AttributeNodes[k].NodeName);
                                    DataLine.Add(N5.AttributeNodes[m].Text);
                                  end;
                              end;
                          end;
                        end;
                    end;

                  //Build Data String from DataLine List
                  for aa := 0 to DataLine.Count - 1 do
                    begin
                      DataString := DataString + DataLine[aa] + Delim;
                    end;

                  //Add DataString to the Output List and clear the DataLine list (at TIER 2)
                  DataOutput.Add(DataString);
                  DataString := '';
                  DataLine.Clear;
                end;
            end;

          //Output the Data to a file here after building the file names to be saved as...
          XMLFileName := TPath.GetFileNameWithoutExtension(in_file);
          CSVFilename := XMLFilename + '.csv';
          LOGFilename := XMLFilename + '_LOG.txt';

          //Check for, or make, the output directory
          if not DirectoryExists(OutDIR) then
            begin
              CreateDir(OutDIR);
              OutLOG.Add('Created new Directory ' + OutDIR);
            end;

          //Assign/Create Output file name
          out_file := TPath.Combine(OutDir, CSVFilename);
          OutLOG.Add('Created new output CSV File: ' + out_file);

          //Output the Data List to TEXT/CSV File
          DataOutput.SaveToFile(out_file);
          OutLOG.Add('Output file SAVED As: ' + out_file);

          //Record Processing time for the LOG
          EndTime := Now;
          DecodeTime(EndTime - StartTime, Hr, Mn, Sec, mSec);
          OutLOG.Add('XML File: ' + in_file + ' Processed in ' + IntToStr(Hr) + ':' + IntToStr(Mn) + ':' + IntToStr(Sec) + '.' + IntToStr(mSec));

          //Output the LOG File
          OutLOG.Add('XML Parsing completed at ' + DateTimeToStr(Now));
          OutLOG.SaveToFile(TPath.Combine(OutDir, LOGFilename));
        end
      else
        begin
          //This is accessed when the given file cannot be accessed
          OutLOG.Add('Input XML File [' + in_file + '] could not be found or accessed!');
        end;

    except
      on E:Exception do
        begin
          Writeln('ERROR: Could not Open or Access the XML Document, because ' + E.Message);
        end;
    end;

  finally
    //Clear Memory
    CoUninitialize;
    DataLine.Free;
    DataOutput.Free;
    OutLOG.Free;
    XML1.Free;
  end;
end;
//_______________________________________________________________________________________________________________________________________________






//**********************  MAIN PROGRAM EXECUTION ************************************************************************************************
var
  Filename : String;
  OutFileName : String;
  Delim : String;
  AOWner : TComponent;
  i, WaitTime : Integer;
  strNote : String;

const
  C : Char = '.';
begin
  try
    //Assign the standardised parameters we expect to receive:
    Filename := ParamStr(1);
    Delim := ParamStr(2);

    WriteLn('XML to CSV Converter Program started at: ' + DateTimeToStr(Now) + '. Received ' + IntToStr(ParamCount) + ' Parameters...');

    //List parameters received externally
    for i := 0 to ParamCount - 1 do
      begin
        WriteLn('Parameter: ' + ParamStr(i));
      end;

    //Check if the File Name passed into the program is valid, exit otherwise
    if (Filename = '') then
      begin
        Writeln('ERROR: NO FILE NAME/PATH WAS SPECIFIED FOR THE XML INPUT FILE!');
        Exit;
      end;

    //Check if we received a VALID colum delimiter, default pipe (|) if not provided by the Caller
    if (Delim = '') then
      begin
        Writeln('Using the DEFAULT Delimiter:   |');
        Delim := '|';
      end;

    //See if the Filename is a valid file and RUN if it is, exit otherwise
    if FileExists(FileName) then
      begin
        Writeln('Process: SEARCHING FOR NEW XML FILE...');
        Sleep(2000);
        AOwner := TComponent.Create(nil);
        ParseXMLData(Filename, OutFilename, AOwner, Delim);
      end
    else
      begin
        Writeln('Could not locate the XML File at: ' + Filename);
        Sleep(10000);
      end;

  except
    on E: Exception do begin
      Writeln('Program error ocurred: ' + E.ClassName, ': ', E.Message);
    end;
  end;
end.
