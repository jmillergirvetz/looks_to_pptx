# imports modules
from looker_sdk import client, models, error
from pptx import Presentation
import io

# client calls will now automatically authenticate using the
# api3credentials specified in 'looker.ini'
sdk = client.setup("../looker.ini")
looker_api_user = sdk.me()

print(looker_api_user)


# the space containing the set of Looks
target_space = 1748

# the powerpoint template used to create the presentation
powerpoint_template = '../looker_template.pptx'



def get_space(space):
    """IN an integer of the space id"""
    """OUT default json of space meta data"""
    return sdk.space(space)

def get_looks(space):
    """IN an integer of the space id"""
    """OUT default: json of all looks"""
    return sdk.space_looks(target_space)

space = get_space(target_space)
looks = get_looks(target_space)



        
def get_number_of_slides(slides):
    """IN pptx presentation slides object"""
    """OUT number of slides"""
    cnt = 0
    for slide in slides:
        cnt += 1
    return cnt
        
        
powerpoint_template = '../presentation_template.pptx'

def get_pres_temp_layout(pptx_temp_path):
    """IN str path to powerpoint template .pptx"""
    """OUT prints out the different layout types then returns powerpoint presentation object"""
    pptx = Presentation(pptx_temp_path)
    for idx, layout in enumerate(pptx.slide_layouts):
        print(idx, layout.name)
    return pptx

pptx = get_pres_temp_layout(powerpoint_template)

title_only_slide = pptx.slide_layouts[5]

def get_look_result(look_id, result_format='png'):
    """IN look's id and the result format; default set to 'png' """
    """OUT result object; default to a 'png' image which returns a bytes"""
    return sdk.run_look(look_id, result_format)

look_request_test = {
        "look_id": looks[0].id, 
        "result_format": 'png', 
        "image_width": 960, 
        "image_height": 540
        }

img = get_look_result(looks[0].id)
print(img)

print(looks[0].id)
shapes = pptx.slides[0].shapes
print(shapes)

# handles converting image bytes string into temp file that isn't written to disk and is constructed as an in-memory binary stream
tmpFile = io.BytesIO(img)

shapes.add_picture(tmpFile, look_request_test['image_width'], look_request_test['image_height'] )

# print(pic)


pptx.save('../out_pres_ex.pptx')





def looks_to_ppt(looks, pptx_temp_path, pptx_output_path):
    """IN json response of all looks in a space | str of powerpoint path | str of desired output path"""
    """OUT creates a pptx presentation with a Look png per slide"""
    
    pptx = get_pres_temp_layout(pptx_temp_path)

    for idx, look in enumerate(looks):
        print(idx, look.id, look.title)

        look_request = {
        "look_id": look.id, 
        "result_format": 'png', 
        "image_width": 960, 
        "image_height": 540
        }

        shapes = pptx.slides[idx].shapes

        shapes.add_picture(get_look_result(look_request))

        


#>>> shapes = Presentation(...).slides[0].shapes

