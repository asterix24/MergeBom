 /* MergeBom is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * Copyright 2015 Daniele Basile <asterix24@gmail.com>
 */

function ExeMergeBom() {
    var mergebomExe= "mergebom_altium.exe";
    //var mergebomExe= "mergebom.bat";
    var wk = GetWorkspace;
    Client.StartServer('WRK');
    if (Client) {
        if (wk != Null) {
            for (i = 0; i < wk.DM_ProjectCount; i++) {
                var p = wk.DM_Projects(i);
                for (j = 0; j < p.DM_GeneratedDocumentCount; j++) {
                    var d = p.DM_GeneratedDocuments(j)
                    var ext = ExtractFileExt(d.DM_FullPath);
                    var bom_file = ExtractFileName(d.DM_FullPath);
                    var path = ExtractFilePath(d.DM_FullPath);

                    bom_file = bom_file.replace(ext,'');

                    if ((ext == '.xls') || (ext == '.xlsx')) {
                        //ShowError(path + " " + d.DM_FullPath);
                        var err = RunApplication(mergebomExe + "  " + d.DM_FullPath);
                        if (err != 0) {
                            ShowError("Errore nell'esecuzione di MergeBom. " + err);
                        } else {
                            var report_doc = Client.OpenDocument('Text', path + '\\' + bom_file + "_report.txt");
                            if (report_doc) {
                                Client.ShowDocument(report_doc);
                                showInfo("MergeBom Done!")
                            } else {
                                ShowError("Errore nell'apertura del documento.");
                            }
                        }
                    }
                }
            }
        }
    }
}

