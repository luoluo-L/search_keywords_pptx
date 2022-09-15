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
        #print(eachfile)
        #print("----------------------")
        for slide_num, slide in enumerate(prs.slides):
            for shape_num, shape in enumerate(slide.shapes):
                if hasattr(shape, "text"):
                    #print(shape.text)

                    if search_term in shape.text.lower():
                        print(('Slide filename: '+str(eachfile)+', slide number: ' + str(slide_num+1) +
                               ', shape number: ' + str(shape_num+1)) )


