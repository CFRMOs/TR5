Imports System.Linq
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.GraphicsInterface

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class PolylineMG
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.GoToEnt = New System.Windows.Forms.Button()
        Me.CopytoLayer = New System.Windows.Forms.Button()
        Me.DividirL = New System.Windows.Forms.Button()
        Me.DeleteENT = New System.Windows.Forms.Button()
        Me.btCFeatureLine = New System.Windows.Forms.Button()
        Me.btConverSelectedEntities = New System.Windows.Forms.Button()
        Me.BtConvertirParcel = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.CoBoxAlignments = New System.Windows.Forms.ComboBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.ComMediciones = New System.Windows.Forms.ComboBox()
        Me.BtmFictureLine = New System.Windows.Forms.Button()
        Me.btAddLine = New System.Windows.Forms.Button()
        Me.BtGetEntity = New System.Windows.Forms.Button()
        Me.BtDataClear = New System.Windows.Forms.Button()
        Me.UpdateData = New System.Windows.Forms.Button()
        Me.IDDatagView = New System.Windows.Forms.Button()
        Me.SetCurrent = New System.Windows.Forms.Button()
        Me.btSetAllLayerColor = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage3.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TextBox3)
        Me.TabPage3.Controls.Add(Me.BtConvertirParcel)
        Me.TabPage3.Controls.Add(Me.btConverSelectedEntities)
        Me.TabPage3.Controls.Add(Me.btCFeatureLine)
        Me.TabPage3.Controls.Add(Me.DeleteENT)
        Me.TabPage3.Controls.Add(Me.DividirL)
        Me.TabPage3.Controls.Add(Me.CopytoLayer)
        Me.TabPage3.Controls.Add(Me.GoToEnt)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(567, 295)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Ediccion"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'GoToEnt
        '
        Me.GoToEnt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GoToEnt.Location = New System.Drawing.Point(6, 6)
        Me.GoToEnt.Name = "GoToEnt"
        Me.GoToEnt.Size = New System.Drawing.Size(129, 43)
        Me.GoToEnt.TabIndex = 3
        Me.GoToEnt.Text = "Zoom Selected Iteam"
        Me.GoToEnt.UseVisualStyleBackColor = True
        '
        'CopytoLayer
        '
        Me.CopytoLayer.Location = New System.Drawing.Point(6, 106)
        Me.CopytoLayer.Name = "CopytoLayer"
        Me.CopytoLayer.Size = New System.Drawing.Size(129, 43)
        Me.CopytoLayer.TabIndex = 4
        Me.CopytoLayer.Text = "Copy to Layer"
        Me.CopytoLayer.UseVisualStyleBackColor = True
        '
        'DividirL
        '
        Me.DividirL.Location = New System.Drawing.Point(276, 6)
        Me.DividirL.Name = "DividirL"
        Me.DividirL.Size = New System.Drawing.Size(129, 43)
        Me.DividirL.TabIndex = 4
        Me.DividirL.Text = "Dividir PL"
        Me.DividirL.UseVisualStyleBackColor = True
        '
        'DeleteENT
        '
        Me.DeleteENT.Location = New System.Drawing.Point(6, 54)
        Me.DeleteENT.Name = "DeleteENT"
        Me.DeleteENT.Size = New System.Drawing.Size(129, 43)
        Me.DeleteENT.TabIndex = 4
        Me.DeleteENT.Text = "Delete Entity"
        Me.DeleteENT.UseVisualStyleBackColor = True
        '
        'btCFeatureLine
        '
        Me.btCFeatureLine.Location = New System.Drawing.Point(141, 6)
        Me.btCFeatureLine.Name = "btCFeatureLine"
        Me.btCFeatureLine.Size = New System.Drawing.Size(129, 43)
        Me.btCFeatureLine.TabIndex = 13
        Me.btCFeatureLine.Text = "Convertir FeatureLine"
        Me.btCFeatureLine.UseVisualStyleBackColor = True
        '
        'btConverSelectedEntities
        '
        Me.btConverSelectedEntities.Location = New System.Drawing.Point(141, 106)
        Me.btConverSelectedEntities.Name = "btConverSelectedEntities"
        Me.btConverSelectedEntities.Size = New System.Drawing.Size(129, 43)
        Me.btConverSelectedEntities.TabIndex = 13
        Me.btConverSelectedEntities.Text = "Convertir FeatureLine"
        Me.btConverSelectedEntities.UseVisualStyleBackColor = True
        '
        'BtConvertirParcel
        '
        Me.BtConvertirParcel.Location = New System.Drawing.Point(141, 55)
        Me.BtConvertirParcel.Name = "BtConvertirParcel"
        Me.BtConvertirParcel.Size = New System.Drawing.Size(129, 43)
        Me.BtConvertirParcel.TabIndex = 14
        Me.BtConvertirParcel.Text = "Convertir Parcel"
        Me.BtConvertirParcel.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(276, 78)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(142, 20)
        Me.TextBox3.TabIndex = 16
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btSetAllLayerColor)
        Me.TabPage1.Controls.Add(Me.SetCurrent)
        Me.TabPage1.Controls.Add(Me.IDDatagView)
        Me.TabPage1.Controls.Add(Me.UpdateData)
        Me.TabPage1.Controls.Add(Me.BtDataClear)
        Me.TabPage1.Controls.Add(Me.BtGetEntity)
        Me.TabPage1.Controls.Add(Me.btAddLine)
        Me.TabPage1.Controls.Add(Me.BtmFictureLine)
        Me.TabPage1.Controls.Add(Me.ComMediciones)
        Me.TabPage1.Controls.Add(Me.ComboBox2)
        Me.TabPage1.Controls.Add(Me.CoBoxAlignments)
        Me.TabPage1.Controls.Add(Me.ComboBox1)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Button2)
        Me.TabPage1.Controls.Add(Me.Button5)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.TextBox1)
        Me.TabPage1.Controls.Add(Me.TextBox2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(567, 295)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Datos Generales"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(28, 262)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(127, 20)
        Me.TextBox2.TabIndex = 7
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(28, 219)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(127, 20)
        Me.TextBox1.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(25, 204)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Station "
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(106, 158)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(89, 43)
        Me.Button5.TabIndex = 4
        Me.Button5.Text = "Mover to Layer"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(17, 158)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(89, 43)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Mover to Layer"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(25, 247)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(54, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "TabName"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Mediciones"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Alineamientos"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(17, 122)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(181, 21)
        Me.ComboBox1.TabIndex = 6
        '
        'CoBoxAlignments
        '
        Me.CoBoxAlignments.FormattingEnabled = True
        Me.CoBoxAlignments.Location = New System.Drawing.Point(14, 21)
        Me.CoBoxAlignments.Name = "CoBoxAlignments"
        Me.CoBoxAlignments.Size = New System.Drawing.Size(181, 21)
        Me.CoBoxAlignments.TabIndex = 6
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(17, 96)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(181, 21)
        Me.ComboBox2.TabIndex = 6
        '
        'ComMediciones
        '
        Me.ComMediciones.FormattingEnabled = True
        Me.ComMediciones.Location = New System.Drawing.Point(17, 69)
        Me.ComMediciones.Name = "ComMediciones"
        Me.ComMediciones.Size = New System.Drawing.Size(181, 21)
        Me.ComMediciones.TabIndex = 6
        '
        'BtmFictureLine
        '
        Me.BtmFictureLine.Location = New System.Drawing.Point(201, 71)
        Me.BtmFictureLine.Name = "BtmFictureLine"
        Me.BtmFictureLine.Size = New System.Drawing.Size(129, 43)
        Me.BtmFictureLine.TabIndex = 11
        Me.BtmFictureLine.Text = "Seleccionar Ficture Line"
        Me.BtmFictureLine.UseVisualStyleBackColor = True
        '
        'btAddLine
        '
        Me.btAddLine.Location = New System.Drawing.Point(201, 120)
        Me.btAddLine.Name = "btAddLine"
        Me.btAddLine.Size = New System.Drawing.Size(129, 43)
        Me.btAddLine.TabIndex = 10
        Me.btAddLine.Text = "Add Lines"
        Me.btAddLine.UseVisualStyleBackColor = True
        '
        'BtGetEntity
        '
        Me.BtGetEntity.Location = New System.Drawing.Point(201, 21)
        Me.BtGetEntity.Name = "BtGetEntity"
        Me.BtGetEntity.Size = New System.Drawing.Size(129, 43)
        Me.BtGetEntity.TabIndex = 9
        Me.BtGetEntity.Text = "Seleccionar un elemento"
        Me.BtGetEntity.UseVisualStyleBackColor = True
        '
        'BtDataClear
        '
        Me.BtDataClear.Location = New System.Drawing.Point(336, 69)
        Me.BtDataClear.Name = "BtDataClear"
        Me.BtDataClear.Size = New System.Drawing.Size(129, 43)
        Me.BtDataClear.TabIndex = 14
        Me.BtDataClear.Text = "Limpiar datos"
        Me.BtDataClear.UseVisualStyleBackColor = True
        '
        'UpdateData
        '
        Me.UpdateData.Location = New System.Drawing.Point(336, 21)
        Me.UpdateData.Name = "UpdateData"
        Me.UpdateData.Size = New System.Drawing.Size(129, 43)
        Me.UpdateData.TabIndex = 13
        Me.UpdateData.Text = "Update"
        Me.UpdateData.UseVisualStyleBackColor = True
        '
        'IDDatagView
        '
        Me.IDDatagView.Location = New System.Drawing.Point(336, 118)
        Me.IDDatagView.Name = "IDDatagView"
        Me.IDDatagView.Size = New System.Drawing.Size(129, 43)
        Me.IDDatagView.TabIndex = 12
        Me.IDDatagView.Text = "Identificar en Data"
        Me.IDDatagView.UseVisualStyleBackColor = True
        '
        'SetCurrent
        '
        Me.SetCurrent.Location = New System.Drawing.Point(201, 169)
        Me.SetCurrent.Name = "SetCurrent"
        Me.SetCurrent.Size = New System.Drawing.Size(129, 43)
        Me.SetCurrent.TabIndex = 15
        Me.SetCurrent.Text = "Set Current"
        Me.SetCurrent.UseVisualStyleBackColor = True
        '
        'btSetAllLayerColor
        '
        Me.btSetAllLayerColor.Location = New System.Drawing.Point(201, 218)
        Me.btSetAllLayerColor.Name = "btSetAllLayerColor"
        Me.btSetAllLayerColor.Size = New System.Drawing.Size(129, 43)
        Me.btSetAllLayerColor.TabIndex = 15
        Me.btSetAllLayerColor.Text = "Set Layers Colors"
        Me.btSetAllLayerColor.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(606, 448)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(575, 321)
        Me.TabControl1.TabIndex = 9
        '
        'PolylineMG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1193, 781)
        Me.Controls.Add(Me.TabControl1)
        Me.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.Name = "PolylineMG"
        Me.Text = "      "
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents BtConvertirParcel As Button
    Friend WithEvents btConverSelectedEntities As Button
    Friend WithEvents btCFeatureLine As Button
    Friend WithEvents DeleteENT As Button
    Friend WithEvents DividirL As Button
    Friend WithEvents CopytoLayer As Button
    Friend WithEvents GoToEnt As Button
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents btSetAllLayerColor As Button
    Friend WithEvents SetCurrent As Button
    Friend WithEvents IDDatagView As Button
    Friend WithEvents UpdateData As Button
    Friend WithEvents BtDataClear As Button
    Friend WithEvents BtGetEntity As Button
    Friend WithEvents btAddLine As Button
    Friend WithEvents BtmFictureLine As Button
    Friend WithEvents ComMediciones As ComboBox
    Friend WithEvents ComboBox2 As ComboBox
    Friend WithEvents CoBoxAlignments As ComboBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TabControl1 As TabControl
End Class
