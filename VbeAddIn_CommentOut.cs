namespace MyVbeAddIn{
    using System;
    using System.Windows.Forms;     // MessageBoxの表示に必要
    using System.Runtime.InteropServices;
    using Extensibility;
    using Office = Microsoft.Office.Core;
    using VBIDE = Microsoft.Vbe.Interop;
    
    //Guidは要変更
    [ComVisible(true), Guid("F2D6539C-8F17-488D-A176-02DAB959823A"), ProgId("VbeAddIn_CommentOut.Connect")]
    public class Connect : Object, Extensibility.IDTExtensibility2{
        private VBIDE.VBE vbe;
        private Office.CommandBarPopup cmdBarPopup;
        private Office.CommandBar cmdBar;
        private Office.CommandBarButton cmdBtn;
        
        public Connect(){}
        
        public void OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref System.Array custom){
            
            vbe = ((VBIDE.VBE) application);
            
            // VBEのメニューにコマンドを登録
            cmdBar = vbe.CommandBars[1];
            cmdBarPopup = (Office.CommandBarPopup)cmdBar.Controls.Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cmdBarPopup.Caption = "&Comment";
            
            // コメントアウト コマンド
            cmdBtn = (Office.CommandBarButton)cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            cmdBtn.Caption = "&Comment Out";
            //cmdBtn.FaceId = 59;               // アイコン
            //cmdBtn.ShortcutText = "Ctrl+w";   // ショートカットキーは表示されるが使用できない
            cmdBtn.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdCommentOut_Click);
            cmdBtn.Visible = true;
        }
        
        /* 選択行をコメントアウトする */
        public void cmdCommentOut_Click(Office.CommandBarButton ctrl, ref bool cancel){
        
            try{
                int startLine, startColumn, endLine, endColumn;
                vbe.ActiveCodePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                
                string buf;
                for (int i = startLine; i <= endLine; i++) {
                    buf = vbe.ActiveCodePane.CodeModule.get_Lines(i, 1);
                    vbe.ActiveCodePane.CodeModule.ReplaceLine(i, "' " + buf);
                }
                
            }catch (Exception){
            //    MessageBox.Show("Comment Out Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref System.Array custom){
            if(cmdBarPopup != null){cmdBarPopup.Delete();}
        }
        
        public void OnAddInsUpdate(ref System.Array custom){}
        public void OnStartupComplete(ref System.Array custom){}
        public void OnBeginShutdown(ref System.Array custom){}
    }
}
