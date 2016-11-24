    Imports System
    Imports System.Runtime.InteropServices
    Imports Extensibility
    Imports Office = Microsoft.Office.Core
    Imports VBIDE = Microsoft.Vbe.Interop
    
    ' Guidは要変更
    <ComVisible(true), Guid("F2D6539C-8F17-488D-A176-02DAB959823A"), ProgId("VbeAddIn_CommentOut.Connect")>
    Public Class Connect 
        Implements IDTExtensibility2 


        Private vbe As VBIDE.VBE 
        Private cmdBarPopup As Office.CommandBarPopup 
        Private WithEvents cmdCommentOut As Office.CommandBarButton
        Private WithEvents cmdGlobalComment As Office.CommandBarButton
        Private WithEvents cmdPrivateComment As Office.CommandBarButton
        Private WithEvents cmdDateComment As Office.CommandBarButton


        Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

            vbe = application
            
            ' VBEのメニューにコマンドを登録
            Dim cmdBar As Office.CommandBar = vbe.CommandBars(1)
            cmdBarPopup = cmdBar.Controls.Add(Office.MsoControlType.msoControlPopup)
            cmdBarPopup.Caption = "&Comment"


            ' コメントアウト コマンド
            cmdCommentOut = cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton)
            With cmdCommentOut
                .Caption = "&Comment Out"
                .ShortcutText = "'"
                .OnAction = "!<VbeAddIn_CommentOut.Connect>"
                .Visible = True
            End With
            
            ' Globalコメント コマンド
            cmdGlobalComment = cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton)
            With cmdGlobalComment
                .Caption = "&Global Comment"
                .ShortcutText = "'***"
                .OnAction = "!<VbeAddIn_CommentOut.Connect>"
                .Visible = True
            End With

            ' Privateコメント コマンド
            cmdPrivateComment = cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton)
            With cmdPrivateComment
                .Caption = "&Private Comment"
                .ShortcutText = "'---"
                .OnAction = "!<VbeAddIn_CommentOut.Connect>"
                .Visible = True
            End With

            ' 日付コメント コマンド
            cmdDateComment = cmdBarPopup.Controls.Add(Office.MsoControlType.msoControlButton)
            With cmdDateComment
                .Caption = "&Date Comment"
                .ShortcutText = "' yyyy/mm/dd"
                .OnAction = "!<VbeAddIn_CommentOut.Connect>"
                .Visible = True
            End With

        End Sub
        
        Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
            If cmdBarPopup Is Nothing Then cmdBarPopup.Delete()
        End Sub
        
        Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
        End Sub
        Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
        End Sub
        Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
        End Sub


        ' 選択行をコメントアウトする
        Public Sub cmdCommentOut_Click(ctrl As Office.CommandBarButton, ByRef cancel As Boolean) Handles cmdCommentOut.Click
        
            Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
            Call vbe.ActiveCodePane.GetSelection(startLine, startColumn, endLine, endColumn)

            Try
                Dim buf As String
                Dim i As Long
                For i = startLine To endLine
                    buf = vbe.ActiveCodePane.CodeModule.Lines(i, 1)
                    vbe.ActiveCodePane.CodeModule.ReplaceLine(i, "'" + buf)
                Next
            Catch ex As Exception
                Call vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "'")
                'Call MsgBox("Comment Out Error", vbCritical + vbOK, "Error")
            End Try
        
        End Sub

        ' Globalスコープ・プロシージャ用のコメントラインを出力
        Public Sub cmdGlobalComment_Click(ctrl As Office.CommandBarButton, ByRef cancel As Boolean) Handles cmdGlobalComment.Click
        
            Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
            Call vbe.ActiveCodePane.GetSelection(startLine, startColumn, endLine, endColumn)

            Try
                Call vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "'****************************************************************************************************")
            Catch ex As Exception
            End Try
        
        End Sub
        
        ' Privateスコープ・プロシージャ用のコメントラインを出力
        Public Sub cmdPrivateComment_Click(ctrl As Office.CommandBarButton, ByRef cancel As Boolean) Handles cmdPrivateComment.Click
        
            Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
            Call vbe.ActiveCodePane.GetSelection(startLine, startColumn, endLine, endColumn)

            Try
                Call vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "'----------------------------------------------------------------------------------------------------")
            Catch ex As Exception
            End Try
        
        End Sub

        ' 日付コメントを出力
        Public Sub cmdDateComment_Click(ctrl As Office.CommandBarButton, ByRef cancel As Boolean) Handles cmdDateComment.Click
        
            Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
            Call vbe.ActiveCodePane.GetSelection(startLine, startColumn, endLine, endColumn)

            Try
                Dim today As String = Date.Now.ToString("yyyy/MM/dd")
                Dim buf As String = vbe.ActiveCodePane.CodeModule.Lines(startLine, 1)
                If buf <> "" Then
                    Call vbe.ActiveCodePane.CodeModule.ReplaceLine(startLine, buf & "    ' " + today)
                Else
                    Call vbe.ActiveCodePane.CodeModule.InsertLines(startLine, "' " + today)
                End If
            Catch ex As Exception
            End Try
        
        End Sub
        
    End Class
