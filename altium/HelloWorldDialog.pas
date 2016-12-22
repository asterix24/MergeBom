{..............................................................................}
{ Summary A simple hello world message using the Script form.                  }
{ Copyright (c) 2003 by Altium Limited                                         }
{..............................................................................}

{..............................................................................}
Var
  HelloWorldForm: THelloWorld;
  myFile : TextFile;
  text   : string;

{..............................................................................}

{..............................................................................}
Procedure THelloWorldForm.bDisplayClick(Sender: TObject);
Begin
     ErrorCode := RunApplication('C:\Program Files (x86)\Saturn PCB Design\PCB Toolkit V5\PCB Toolkit V5.65.exe');
     If ErrorCode <> 0 Then
        ShowError('System cannot start PCB Toolkit ' + GetErrorMessage(ErrorCode));
     end;

     // Try to open the Test.txt file for writing to
     AssignFile(myFile, 'C:\Users\asterix\src\MergeBom\altium\mergebom_report.txt');
     Reset(myFile);
     // Display the file contents

     while not Eof(myFile) do
     begin
          ReadLn(myFile, text);
           If Not VarIsNull(text) Then
           Begin
                log.Lines.Add(text);
           end;
     end;
    // Close the file for the last time
   CloseFile(myFile);
End;
{..............................................................................}

{..............................................................................}
Procedure THelloWorldForm.bCloseClick(Sender: TObject);
Begin
    close;
End;
{..............................................................................}

{..............................................................................}
Procedure RunHelloWorld;
Begin
     // Get the current directory
     text := GetCurrentDir;
     ShowMessage('Current directory = '+text);
     HelloWorldForm.ShowModal;
End;
{..............................................................................}

{..............................................................................}
End.
