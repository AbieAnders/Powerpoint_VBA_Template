Sub ppt_maker_macro()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim slide1 As Object
    Dim slide2 As Object
    Dim slide3 As Object
    Dim slide4 As Object
    Dim slide5 As Object
    Dim slide6 As Object
    Dim slide7 As Object
    Dim slide8 As Object
    Dim slide9 As Object

    ' Create a PowerPoint application object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' Create a new presentation
    Set pptPresentation = pptApp.Presentations.Add

    ' Slide 1: Introduction
    Set slide1 = pptPresentation.Slides.Add(1, 1) ' ppLayoutTitle
    slide1.Shapes.Title.TextFrame.TextRange.Text = "Unleashing the Power of Python: The Right Choice for Java Veterans" ' Prints the desired text on the selected slide object.

    ' Slide 2: From Java to Python - Why Consider the Shift?
    Set slide2 = pptPresentation.Slides.Add(2, 5) ' ppLayoutTwoObjects
    slide2.Shapes.Title.TextFrame.TextRange.Text = "From Java to Python - Why Consider the Shift?"
    slide2.Shapes(2).TextFrame.TextRange.Text = "Python's concise and readable syntax reduces code verbosity, promotes faster and clearer development compared to Java. Python's dynamic typing eliminates the need for explicit variable declarations, simplifying code maintenance and increasing productivity."
    ' Add image to slide 2
    slide2.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

    ' Slide 3: Python - Easy to Use, Easy to Develop, and Easy to Maintain
    Set slide3 = pptPresentation.Slides.Add(3, 5) ' ppLayoutTwoObjects
    slide3.Shapes.Title.TextFrame.TextRange.Text = "Python - Easy to Use, Easy to Develop, and Easy to Maintain"
    slide3.Shapes(2).TextFrame.TextRange.Text = "Python's extensive standard library and strong community support provide a rich ecosystem for rapid development and seamless integration."
    ' Add image to slide 3
    slide3.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

        ' Slide 4: Python's Scalability and REST API Support
    Set slide4 = pptPresentation.Slides.Add(4, 5) ' ppLayoutTwoObjects
    slide4.Shapes.Title.TextFrame.TextRange.Text = "Python's Scalability and REST API Support"
    slide4.Shapes(2).TextFrame.TextRange.Text = "Dive into Python's scalability prowess! Building RESTful APIs with Python's frameworks propels seamless integration across your systems."
    ' Add image to slide 4
    slide4.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

    ' Slide 5: Case Studies - Real-World Python Success Stories
    Set slide5 = pptPresentation.Slides.Add(5, 5) ' ppLayoutTwoObjects
    slide5.Shapes.Title.TextFrame.TextRange.Text = "Case Studies - Real-World Python Success Stories"
    slide5.Shapes(2).TextFrame.TextRange.Text = "Hear the triumphs of companies that embraced Python. Witness their growth, efficiency, and joy in adopting this versatile language."
    ' Add image to slide 5
    slide5.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

    ' Slide 6: Getting Started with Python
    Set slide6 = pptPresentation.Slides.Add(6, 5) ' ppLayoutTwoObjects
    slide6.Shapes.Title.TextFrame.TextRange.Text = "Getting Started with Python"
    slide6.Shapes(2).TextFrame.TextRange.Text = "Are you ready to embrace Python's power? Let's embark on your Python journey together!"
    ' Add image to slide 6
    slide6.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

    ' Slide 7: Python's Bright Future
    Set slide7 = pptPresentation.Slides.Add(7, 5) ' ppLayoutTwoObjects
    slide7.Shapes.Title.TextFrame.TextRange.Text = "Python's Bright Future"
    slide7.Shapes(2).TextFrame.TextRange.Text = "The future of Python is brighter than ever before! Let's explore its endless possibilities and its role in shaping future technologies."
    ' Add image to slide 7
    slide7.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

    ' Slide 8: Q&A Session
    Set slide8 = pptPresentation.Slides.Add(8, 5) ' ppLayoutTwoObjects
    slide8.Shapes.Title.TextFrame.TextRange.Text = "Q&A Session"
    slide8.Shapes(2).TextFrame.TextRange.Text = "Your questions and curiosity are welcome! Let's have an engaging discussion about Python's incredible potential."
    ' Add image to slide 8
    slide8.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200

    ' Slide 9: Closing - Thank You
    Set slide9 = pptPresentation.Slides.Add(9, 1) ' ppLayoutTitle
    slide9.Shapes.Title.TextFrame.TextRange.Text = "Thank You!"
    ' Add image to slide 9 if necessary(Closing slide)
    ' slide9.Shapes.AddPicture "C:\Users\Dell\Downloads\Lava Cat.jpg", False, True, 550, 200, 300, 200
End Sub
