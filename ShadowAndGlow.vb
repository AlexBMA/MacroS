Sub GlowAndShadow()

Dim vslide As Slide

Dim pre As Presentation
Set pre = ActivePresentation


For Each vslide In ActivePresentation.Slides

    'Set transperancy = vslide.Shapes("Title 1").Glow'
    
     vslide.Shapes("Text Box 7").TextFrame2.TextRange.Font.Glow.Radius = 6
     vslide.Shapes("Text Box 7").TextFrame2.TextRange.Font.Glow.Color.RGB = (0)
     vslide.Shapes("Text Box 7").TextFrame2.TextRange.Font.Glow.Transparency = 0.6
     vslide.Shapes("Text Box 7").TextFrame2.TextRange.Font.Line.Visible = msoCTrue
     vslide.Shapes("Text Box 7").TextFrame2.TextRange.Font.Line.Weight = 2.25
     
     vslide.Shapes("Text Box 8").TextFrame2.TextRange.Font.Glow.Radius = 6
     vslide.Shapes("Text Box 8").TextFrame2.TextRange.Font.Glow.Color.RGB = (0)
     vslide.Shapes("Text Box 8").TextFrame2.TextRange.Font.Glow.Transparency = 0.6
     vslide.Shapes("Text Box 8").TextFrame2.TextRange.Font.Line.Visible = msoCTrue
     vslide.Shapes("Text Box 8").TextFrame2.TextRange.Font.Line.Weight = 2.25
     
    Debug.Print ("abc")
   
    
Next

End Sub
