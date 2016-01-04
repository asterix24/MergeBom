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

function dirname(path) {
     return path.replace(/\\[^\\]*$/g, '');
}

function ExeMergeBom() {
    var mergebomExe= "mergebom_altium.exe";
    //var mergebomExe= "mergebom.bat";
    var wk = GetWorkspace;
    Client.StartServer('WRK');
    if (Client) {
        if (wk != Null) {
            var p = wk.DM_FocusedProject();
            var path = dirname(p.DM_ProjectFullPath);
            if (path) {
                var err = RunApplication(mergebomExe + "  " + path);
                if (err != 0) {
                    ShowMessage("Errore nell'esecuzione di MergeBom. " + err);
                } else {
                    var report_doc = Client.OpenDocument('Text', path + '\\' + "mergebom_report.txt");
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
