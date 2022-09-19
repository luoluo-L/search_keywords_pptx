# search key words in slides

import sys

# To install the package, use the following command:
# pip install python-pptx

from pptx import Presentation
import glob

if __name__ == '__main__':

    #test_pptx_path = "Test Path with pptx slides"
    test_pptx_path = sys.argv[1]
    search_term = sys.argv[2] #'results'

    for eachfile in glob.glob(test_pptx_path):
        prs = Presentation(eachfile)
        #print(eachfile)
        #print("----------------------")
        for slide_num, slide in enumerate(prs.slides):
            for shape_num, shape in enumerate(slide.shapes):
                if hasattr(shape, "text"):

                    if search_term in shape.text.lower():
                        print(('Slide filename: '+str(eachfile)+', slide number: ' + str(slide_num+1) +
                               ', shape number: ' + str(shape_num+1)) )


