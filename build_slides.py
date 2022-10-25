from __future__ import print_function

from googleapiclient.errors import HttpError
from classes import GoogleSlideClient
from constants import SCOPES, PRESENTATION_ID

def add_to_slide(object_id, client: GoogleSlideClient, presentation_id):
    # pylint: disable=maybe-no-member
    try:
        # Add a slide at index 1 using the predefine
        requests = [
            {
                "insertText": {
                    "objectId": object_id,
                    "text": "My List\n\tItem 1\n\t\tItem 2\n\t\t\tItem 3",
                    "insertionIndex": 0
                },
            },
            {
                "createParagraphBullets": {
                    "objectId": object_id,
                    "bulletPreset": "BULLET_ARROW_DIAMOND_DISC",
                    "textRange": {
                        "type": "ALL"
                    }}
            }

        ]

        # If you wish to populate the slide with elements,
        # add element create requests here, using the page_id.

        # Execute the request.
        body = {
            'requests': requests
        }
        response = client.service.presentations() \
            .batchUpdate(presentationId=presentation_id, body=body).execute()
        print(response)

    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response


def create_slide(presentation_id, client: GoogleSlideClient, page_id):
    try:
        # Add a slide at index 1 using the predefine
        requests = [
            {
                'createSlide': {
                    'objectId': page_id,
                    'insertionIndex': '1',
                    'slideLayoutReference': {
                        'predefinedLayout': 'TITLE_AND_BODY'
                    }
                },
            }
        ]

        # If you wish to populate the slide with elements,
        # add element create requests here, using the page_id.

        # Execute the request.
        body = {
            'requests': requests
        }
        response = client.service.presentations() \
            .batchUpdate(presentationId=presentation_id, body=body).execute()
        create_slide_response = response.get('replies')[0].get('createSlide')
        print(f"Created slide with ID:"
              f"{(create_slide_response.get('objectId'))}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response


def build_slides(client: GoogleSlideClient):
    add_to_slide(object_id="SLIDES_API1821489010_1",
                 presentation_id=PRESENTATION_ID, client=client)

    presentation = client.service.presentations().get(
        presentationId=PRESENTATION_ID).execute()
    slides = presentation.get('slides')

    print('The presentation contains {} slides:'.format(len(slides)))
    for slide in slides:
        for element in slide['pageElements']:
            for key, val in element.items():
                print(key, val)


if __name__ == '__main__':
    client = GoogleSlideClient(scopes=SCOPES)
    build_slides(client=client)
