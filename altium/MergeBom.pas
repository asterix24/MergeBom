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
   WorkSpace       : IWorkSpace;
   VersionFileName : TDynamicString;
   MergeBomLog     : TDynamicString;
   PrjFullPath     : TDynamicString;
   Project        : IProject;
   Document       : IDocument;
   I              : Integer;
   item           : String;
   logfile        : TextFile;
   line           : String;
   ErrorCode      : Integer;

procedure TMergeBomForm.bCloseClick(Sender: TObject);
begin
     Close;
end;

Procedure MainMergeBom;
Begin
     MergeBomForm.ShowModal;
End;

procedure TMergeBomForm.brunClick(Sender: TObject);
begin
     log.Text := 'Run MergeBom Script..';

     VersionFileName := Nil;
     MergeBomLog := Nil;
     WorkSpace := GetWorkSpace.DM_FocusedProject;

    If WorkSpace = Nil Then
       begin
            log.Lines.Add('Invalid Workspace');
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
                log.Lines.Add('Found Versione file..');
                log.Lines.Add(PrjFullPath);
                log.Lines.Add(VersionFileName);
           end;
    End;

    If Not VarIsNull(VersionFileName) and fileexists(VersionFileName) Then
       Begin
            ErrorCode := RunApplication('mergebom_altium.exe -d '+PrjFullPath+' -v'+ VersionFileName);
            If ErrorCode <> 0 Then
            Begin
               log.Lines.Add('Unable to exe script to merge bom.');
               log.Lines.Add(GetErrorMessage(ErrorCode));
               Exit;
            end;
            log.Lines.Add('exe..');
            MergeBomLog := PrjFullPath + '\mergebom_report.txt';
            log.Lines.Add(MergeBomLog);

            If Not VarIsNull(MergeBomLog) and fileexists(MergeBomLog) Then
               Begin
                   AssignFile(logfile, MergeBomLog);
                   Reset(logfile);
                   while not Eof(logfile) do
                         begin
                              ReadLn(logfile, line);
                              If Not VarIsNull(line) Then log.Lines.Add(line);
                         end;
                   CloseFile(logfile);
               end
            else
                Begin
                   log.Lines.Add('MergeLogFile not found..');
               end;

       end
    else
        begin
             log.Text := 'No Version File found.';
        end;
    end;
end;

