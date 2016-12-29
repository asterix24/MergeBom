{..............................................................................}
{   MergeBom is free software; you can redistribute it and/or modify           }
{  it under the terms of the GNU General Public License as published by        }
{  the Free Software Foundation; either version 2 of the License, or           }
{  (at your option) any later version.                                         }
{                                                                              }
{  This program is distributed in the hope that it will be useful,             }
{  but WITHOUT ANY WARRANTY; without even the implied warranty of              }
{  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the               }
{  GNU General Public License for more details.                                }
{                                                                              }
{  You should have received a copy of the GNU General Public License           }
{  along with this program; if not, write to the Free Software                 }
{  Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA  }
{                                                                              }
{  Copyright 2015 Daniele Basile <asterix24@gmail.com>                         }
{                                                                              }
{..............................................................................}

Var
   MergeBomForm: TMergeBom;
   mergelog        : string;
   myFile          : TextFile;
   text            : string;
   WorkSpace       : IWorkSpace;
   VersionFileName : TDynamicString;
   PrjFullPath     : TDynamicString;

   Project        : IProject;

   K              : Integer;
   ReportFile     : Text;
   ReportDocument : IServerDocument;


Procedure MainMergeBom;
Begin
     MergeBomForm.ShowModal;
End;

procedure TMergeBomForm.closeClick(Sender: TObject);
begin
     close;
end;

procedure TMergeBomForm.runClick(Sender: TObject);
var
   Document       : IDocument;
   I              : Integer;
   item           : String;
begin
    VersionFileName := Nil;
    WorkSpace := GetWorkSpace.DM_FocusedProject;

    If WorkSpace = Nil Then
       begin
            log.Text := 'Invalid Workspace';
            Exit;
       end;

    For i := 0 To WorkSpace.DM_LogicalDocumentCount - 1 Do
    Begin
        Document := WorkSpace.DM_LogicalDocuments(i);
        item := ExtractFileName(Document.DM_FullPath);
        PrjFullPath := ExtractFilePath(Document.DM_FullPath);
        if ansicomparestr(item, 'version.txt') = 0 then
           begin
                VersionFileName := Document.DM_FullPath;
                log.Lines.Add(PrjFullPath);
                log.Lines.Add(VersionFileName);
           end;
    End;

    If Not VarIsNull(text) Then
       Begin
         ShowMessage('Run Cmd');
          {ErrorCode := RunApplication('C:\Program Files (x86)\Saturn PCB Design\PCB Toolkit V5\PCB Toolkit V5.65.exe');
     If ErrorCode <> 0 Then
        ShowError('System cannot start PCB Toolkit ' + GetErrorMessage(ErrorCode));
     end;}
       end
       else
           begin
                log.Text := 'No Version File found.';
           end
    end;
end;

procedure TMergeBomForm.showlogClick(Sender: TObject);
begin
     mergelog := 'Z:\src\fues-sch\seb3\mergebom_report.txt';
     If Not VarIsNull(VersionFileName) Then
        Begin
        log.Text := 'No Version File found.';
        end;

     // Try to open the Test.txt file for writing to
     if fileexists(VersionFileName) then
     begin
          AssignFile(myFile, VersionFileName);
          Reset(myFile);
          while not Eof(myFile) do
          begin
               ReadLn(myFile, text);
               If Not VarIsNull(text) Then
               Begin
                    log.Lines.Add(text);
               end;
          end;
          CloseFile(myFile);
     end
     else
         log.Text := 'MergeBom report not found.';
     end;
end;

