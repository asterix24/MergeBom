function dirname(path) {
     return path.replace(/\\/g,'/').replace(/\/[^\/]*$/, '');
}

function ExeMergeBom() {
    var wk = GetWorkspace;
    Client.StartServer('WRK');
    if (Client) {
        if (wk != Null) {
            var p = wk.DM_FocusedProject();
            var path = dirname(p.DM_ProjectFullPath);
            if (path) {
                var err = RunApplication('mergebom.bat' + "  " + path);
                if (err != 0) {
                    ShowMessage("Errore nell'esecuzione di MergeBom. " + err);
                } else {
                    var report_doc = Client.OpenDocument('Text', path + "\\" + "mergebom_report.txt");
                    if (report_doc) {
                        Client.ShowDocument(report_doc);
                    } else {
                        ShowMessage("Errore nell'apertura del documento.");
                    }
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
   //for (i = 0; i < wk.DM_ProjectCount; i++) {
//var p = wk.DM_Projects(i);
//}

}

