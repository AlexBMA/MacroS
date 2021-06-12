Option Explicit

Sub PPT_demo()

Dim pre As Presentation
Set pre = ActivePresentation


ShapeUpdate
Slide

Set pre = Nothing

End Sub

Function db(msg As String)
Debug.Print msg


End Function

Function ShapeUpdate()
Dim vslide As Slide

For Each vslide In ActivePresentation.Slides
        
    Set rect = vslide.Shapes("Rectangle 3")
    rect.Delete
    
    Set rect = vslide.Shapes("Rectangle 4")
    rect.Delete
Next


End Function

Function Slide()

Dim vslide As Slide

Dim s As Shape

Dim rect As Shape


ActivePresentation.SlideMaster.Background.Fill.ForeColor.RGB = RGB(0, 0, 0)
ActivePresentation.SlideMaster.Background.Fill.BackColor.RGB = RGB(0, 0, 0)


Debug.Print "abcd"


For Each vslide In ActivePresentation.Slides


    'Set rect = vslide.Shapes("Rectangle 3")
    'If Not rect Is Nothing Then
    '    rect.Delete
    'End If
    
    'Set rect = vslide.Shapes("Rectangle 4")
    'rect.Delete
     
    'Debug.Print vslide.Background.Fill.BackColor.RGB
    
    'vslide.Background.Fill.BackColor.RGB = RGB(0, 0, 0)
     
    'Debug.Print vslide.Background.Fill.BackColor.RGB
     
     
     
    For Each s In vslide.Shapes
    
    
    'Debug.Print (s.MediaType)
    
    Debug.Print (s.Name)
    
    If s.HasTextFrame Then

        With s.TextFrame

            If .HasText Then
                
                'Debug.Print .TextRange.Text
            
                If .TextRange.Text = "O, ce valuri, de-ndurare" Then
                
                    .TextRange.Delete
                    
                    Debug.Print "Gasit si sters"
                    
                
                End If
                
            
            End If

        End With
        
    End If
    
   
        
    
    Next

Next






db ("here")


End Function







