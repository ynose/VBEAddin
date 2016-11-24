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
/*          private Office.CommandBar cmdBar;
          private Office.CommandBarButton cmdCommentOut;
          private Office.CommandBarButton cmdGlobalComment;
          private Office.CommandBarButton cmdPrivateComment;
*/
        
        public Connect(){}
        
        public void OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref System.Array custom){
            
            vbe = ((VBIDE.VBE) application);
            
            // VBEのメニューにコマンドを登録
            Office.CommandBar cmdBar = vbe.CommandBars[1];
            
            cmdBarPopup = (Office.CommandBarPopup)cmdBar.Controls.Add(Office.MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            cmdBarPopup.Caption = "&Comment";

            
            // コメントアウト コマンド
            Office.CommandBarButton cmdCommentOut;
            cmdCommentOut = (Office.CommandBarButton)cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            cmdCommentOut.Caption = "&Comment Out";
            cmdCommentOut.ShortcutText = "'";
            cmdCommentOut.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdCommentOut_Click);
            cmdCommentOut.Visible = true;

            // Globalコメント コマンド
            Office.CommandBarButton cmdGlobalComment;
            cmdGlobalComment = (Office.CommandBarButton)cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            cmdGlobalComment.Caption = "&Global Comment";
            cmdGlobalComment.ShortcutText = "'***";
            cmdGlobalComment.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdGlobalComment_Click);
            cmdGlobalComment.Visible = true;

            // Privateコメント コマンド
            Office.CommandBarButton cmdPrivateComment;
            cmdPrivateComment = (Office.CommandBarButton)cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            cmdPrivateComment.Caption = "&Private Comment";
            cmdPrivateComment.ShortcutText = "'---";
            cmdPrivateComment.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdPrivateComment_Click);
            cmdPrivateComment.Visible = true;

            // 日付コメント コマンド
            Office.CommandBarButton cmdDateComment;
            cmdDateComment = (Office.CommandBarButton)cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            cmdDateComment.Caption = "&Date Comment";
            cmdDateComment.ShortcutText = "' yyyy/mm/dd";
            cmdDateComment.Click += new Office._CommandBarButtonEvents_ClickEventHandler(cmdDateComment_Click);
            cmdDateComment.Visible = true;
        }
        
        /* 選択行をコメントアウトする */
        public void cmdCommentOut_Click(Office.CommandBarButton ctrl, ref bool cancel){
        
            try{
                int startLine, startColumn, endLine, endColumn;
                vbe.ActiveCodePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                
                string buf;
                for (int i = startLine; i <= endLine; i++) {
                    buf = vbe.ActiveCodePane.CodeModule.get_Lines(i, 1);
                    vbe.ActiveCodePane.CodeModule.ReplaceLine(i, "'" + buf);
                }
                
            }catch (Exception){
            //    MessageBox.Show("Comment Out Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /* Globalスコープ・プロシージャ用のコメントラインを出力 */
        public void cmdGlobalComment_Click(Office.CommandBarButton ctrl, ref bool cancel){
        
            try{
                int startLine, startColumn, endLine, endColumn;
                vbe.ActiveCodePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "'****************************************************************************************************");
            }catch (Exception){
            }
        }

        /* Privateスコープ・プロシージャ用のコメントラインを出力 */
        public void cmdPrivateComment_Click(Office.CommandBarButton ctrl, ref bool cancel){
        
            try{
                int startLine, startColumn, endLine, endColumn;
                vbe.ActiveCodePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "'----------------------------------------------------------------------------------------------------");
            }catch (Exception){
            }
        }

        /* 日付コメントを出力 */
        public void cmdDateComment_Click(Office.CommandBarButton ctrl, ref bool cancel){
        
            try{
                int startLine, startColumn, endLine, endColumn;
                vbe.ActiveCodePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);

                string now = DateTime.Now.ToString("yyyy/MM/dd");
                string buf = vbe.ActiveCodePane.CodeModule.get_Lines(startLine, 1);
                if (String.IsNullOrEmpty(buf) == false ) {
                    vbe.ActiveCodePane.CodeModule.ReplaceLine(startLine, buf + "    ' " + now);
                } else {
                    vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "' " + now);
                }

            }catch (Exception){
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
