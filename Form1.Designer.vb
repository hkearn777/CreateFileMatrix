<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
    lbModels = New ListBox()
    btnCreateFileMatrix = New Button()
    btnClose = New Button()
    Label1 = New Label()
    txtSandboxRoot = New TextBox()
    Label2 = New Label()
    btnFindModels = New Button()
    lblModelsFound = New Label()
    lblModelsProcessed = New Label()
    lblFileMatrix = New Label()
    Label3 = New Label()
    lbMessages = New ListBox()
    SuspendLayout()
    ' 
    ' lbModels
    ' 
    lbModels.FormattingEnabled = True
    lbModels.ItemHeight = 25
    lbModels.Location = New Point(74, 88)
    lbModels.Name = "lbModels"
    lbModels.ScrollAlwaysVisible = True
    lbModels.SelectionMode = SelectionMode.MultiExtended
    lbModels.Size = New Size(1467, 254)
    lbModels.TabIndex = 0
    ' 
    ' btnCreateFileMatrix
    ' 
    btnCreateFileMatrix.Location = New Point(1195, 517)
    btnCreateFileMatrix.Name = "btnCreateFileMatrix"
    btnCreateFileMatrix.Size = New Size(164, 34)
    btnCreateFileMatrix.TabIndex = 1
    btnCreateFileMatrix.Text = "Create File Matrix"
    btnCreateFileMatrix.UseVisualStyleBackColor = True
    ' 
    ' btnClose
    ' 
    btnClose.Location = New Point(1429, 517)
    btnClose.Name = "btnClose"
    btnClose.Size = New Size(112, 34)
    btnClose.TabIndex = 2
    btnClose.Text = "Close"
    btnClose.UseVisualStyleBackColor = True
    ' 
    ' Label1
    ' 
    Label1.AutoSize = True
    Label1.Location = New Point(74, 14)
    Label1.Name = "Label1"
    Label1.Size = New Size(125, 25)
    Label1.TabIndex = 3
    Label1.Text = "Sandbox Root"
    ' 
    ' txtSandboxRoot
    ' 
    txtSandboxRoot.Location = New Point(205, 11)
    txtSandboxRoot.Name = "txtSandboxRoot"
    txtSandboxRoot.Size = New Size(940, 31)
    txtSandboxRoot.TabIndex = 4
    ' 
    ' Label2
    ' 
    Label2.AutoSize = True
    Label2.Font = New Font("Segoe UI", 9F, FontStyle.Bold)
    Label2.Location = New Point(74, 57)
    Label2.Name = "Label2"
    Label2.Size = New Size(136, 25)
    Label2.TabIndex = 5
    Label2.Text = "List of Models:"
    ' 
    ' btnFindModels
    ' 
    btnFindModels.Location = New Point(1166, 9)
    btnFindModels.Name = "btnFindModels"
    btnFindModels.Size = New Size(153, 34)
    btnFindModels.TabIndex = 6
    btnFindModels.Text = "Find Models"
    btnFindModels.UseVisualStyleBackColor = True
    ' 
    ' lblModelsFound
    ' 
    lblModelsFound.AutoSize = True
    lblModelsFound.Location = New Point(74, 485)
    lblModelsFound.Name = "lblModelsFound"
    lblModelsFound.Size = New Size(143, 25)
    lblModelsFound.TabIndex = 7
    lblModelsFound.Text = "Models found: 0"
    ' 
    ' lblModelsProcessed
    ' 
    lblModelsProcessed.AutoSize = True
    lblModelsProcessed.Location = New Point(74, 512)
    lblModelsProcessed.Name = "lblModelsProcessed"
    lblModelsProcessed.Size = New Size(171, 25)
    lblModelsProcessed.TabIndex = 8
    lblModelsProcessed.Text = "Models processed:0"
    ' 
    ' lblFileMatrix
    ' 
    lblFileMatrix.AutoSize = True
    lblFileMatrix.Location = New Point(74, 540)
    lblFileMatrix.Name = "lblFileMatrix"
    lblFileMatrix.Size = New Size(101, 25)
    lblFileMatrix.TabIndex = 9
    lblFileMatrix.Text = "File Matrix: "
    ' 
    ' Label3
    ' 
    Label3.AutoSize = True
    Label3.Font = New Font("Segoe UI", 9F, FontStyle.Bold)
    Label3.Location = New Point(76, 359)
    Label3.Name = "Label3"
    Label3.Size = New Size(151, 25)
    Label3.TabIndex = 10
    Label3.Text = "List of Messages"
    ' 
    ' lbMessages
    ' 
    lbMessages.FormattingEnabled = True
    lbMessages.ItemHeight = 25
    lbMessages.Location = New Point(77, 386)
    lbMessages.Name = "lbMessages"
    lbMessages.ScrollAlwaysVisible = True
    lbMessages.Size = New Size(1464, 79)
    lbMessages.TabIndex = 11
    ' 
    ' Form1
    ' 
    AutoScaleDimensions = New SizeF(10F, 25F)
    AutoScaleMode = AutoScaleMode.Font
    ClientSize = New Size(1635, 585)
    Controls.Add(lbMessages)
    Controls.Add(Label3)
    Controls.Add(lblFileMatrix)
    Controls.Add(lblModelsProcessed)
    Controls.Add(lblModelsFound)
    Controls.Add(btnFindModels)
    Controls.Add(Label2)
    Controls.Add(txtSandboxRoot)
    Controls.Add(Label1)
    Controls.Add(btnClose)
    Controls.Add(btnCreateFileMatrix)
    Controls.Add(lbModels)
    Name = "Form1"
    Text = "Form1"
    ResumeLayout(False)
    PerformLayout()
  End Sub

  Friend WithEvents lbModels As ListBox
  Friend WithEvents btnCreateFileMatrix As Button
  Friend WithEvents btnClose As Button
  Friend WithEvents Label1 As Label
  Friend WithEvents txtSandboxRoot As TextBox
  Friend WithEvents Label2 As Label
  Friend WithEvents btnFindModels As Button
  Friend WithEvents lblModelsFound As Label
  Friend WithEvents lblModelsProcessed As Label
  Friend WithEvents lblFileMatrix As Label
  Friend WithEvents Label3 As Label
  Friend WithEvents lbMessages As ListBox

End Class
