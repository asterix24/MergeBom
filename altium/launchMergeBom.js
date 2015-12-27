function ExeMergeBom() {
  var wk = GetWorkspace;
  if (wk != Null) {
     for (i = 0; i < wk.DM_ProjectCount; i++) {
         var p = wk.DM_Projects(i);
         var path = p.DM_ProjectFullPath
         if (path) {
            var err = RunApplication('mergebom.bat' + "  " + path);
            if (err != 0) {
               ShowMessage("Errore nell'esecuzione di MergeBom. " + err);
            } else {
               ShowMessage("Done");
            }
         }
     }
  }
}

function Prova() {
   var w = GetWorkspace();
   var f = OpenFile("cioa");
   showMessage("ciao");
   /*
   if (Client) {
      var report_doc = Client.OpenDocument('Text', "prova");
      if (report_doc) {
         Client.ShowDocument(report_doc);
      }
   }
   */
}

