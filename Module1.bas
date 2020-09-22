Attribute VB_Name = "Module1"


Sub AlphaBlendPictures(ByVal BlendValue As Integer)

With BF
    .BlendOp = AC_SRC_OVER
    .BlendFlags = 0
    .SourceConstantAlpha = BlendValue
    .AlphaFormat = 0
End With
    
RtlMoveMemory lBF, BF, 4 'Convert the BLENDFUNCTION-structure to a Long
 
XCoord = Form1.FadeAnswerPictureBox.Left
YCoord = Form1.FadeAnswerPictureBox.Top

ThisWidth = Form1.FadeAnswerPictureBox.ScaleWidth
ThisHeight = Form1.FadeAnswerPictureBox.ScaleHeight

Form1.FadeAnswerPictureBox.Visible = False
 
 
'AlphaBlend the picture from Picture1 over the picture of Picture2
  
AlphaBlend Form1.FadeAnswerPictureBox.hdc, 0, 0, ThisWidth, ThisHeight, Form1.hdc, XCoord, YCoord, ThisWidth, ThisHeight, lBF
Form1.FadeAnswerPictureBox.Visible = True

Form1.FadeAnswerPictureBox.Refresh

End Sub
