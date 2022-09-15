# search key words in slides

if __name__ == '__main__':

    test_pptx_path = "Test Path with pptx slides"

    #To install the package, use the following command:
    #pip install python-pptx 
    from pptx import Presentation
    import glob

    search_term = 'results'
    for eachfile in glob.glob(test_pptx_path):
        prs = Presentation(eachfile)
        print(eachfile)
        print("----------------------")
        for slide_num, slide in enumerate(prs.slides):
            for shape_num, shape in enumerate(slide.shapes):
                if hasattr(shape, "text"):
                    #print(shape.text)

                    if search_term in shape.text.lower():
                        print(('Slide filename: '+str(eachfile)+', slide number: ' + str(slide_num+1) +
                               ', shape number: ' + str(shape_num+1)) )



    """
    import aspose.slides as slides  -- it appears to need a licience

    # Get all the text from presentation
    text = slides.PresentationFactory().get_presentation_text(test_pptx_path,
                                                              slides.TextExtractionArrangingMode.UNARRANGED)

    # Load the presentation to get slide count
    with slides.Presentation(test_pptx_path) as ppt:

        # Loop through slides in the presentation
        for index in range(ppt.slides.length):
            # Print text of desired sections such as slide's text, layout text, notes, etc.
            print(text.slides_text[index].text)
            print(text.slides_text[index].layout_text)
            print(text.slides_text[index].master_text)
            print(text.slides_text[index].notes_text)
    """




