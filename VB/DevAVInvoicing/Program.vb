Imports System
Imports System.Windows.Forms

Namespace DevAVInvoicing

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread>
        Sub Main()
            DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle("Office 2016 Colorful")
            DevExpress.Skins.SkinManager.EnableFormSkins()
            Call Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Call Application.Run(New Form1())
        End Sub
    End Module
End Namespace
