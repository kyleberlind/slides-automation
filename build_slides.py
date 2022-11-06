from __future__ import print_function

from typing import List, Dict
from googleapiclient.errors import HttpError
from classes import GoogleClient
from constants import SCOPES, PRESENTATION_ID
from datetime import datetime


def add_to_slide(object_id, client: GoogleClient, presentation_id):
    try:
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
        response = update_presetation_with_requests(
            presentation_id=presentation_id, client=client, requests=requests)
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Slides not created")
        return error

    return response


def create_slide(presentation_id, client: GoogleClient, page_id):
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
        response = update_presetation_with_requests(
            presentation_id=presentation_id, client=client, requests=requests)
        create_slide_response = response.get('replies')[0].get('createSlide')
        print(f"Created slide with ID:"
              f"{(create_slide_response.get('objectId'))}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

    return response


def build_slides(client: GoogleClient):
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


def update_presetation_with_requests(presentation_id, requests: List, client: GoogleClient):
    try:
        body = {
            'requests': requests
        }
        response = client.slide_service.presentations() \
            .batchUpdate(presentationId=presentation_id, body=body).execute()
        return response
    except HttpError as error:
        print(f"An error occurred: {error}")
        raise error


def create_presentation_from_template(customer_name: str, template_presentation_id: str, client: GoogleClient, template_data: Dict):
    """
    creates a presentation from a template
    """

    copy_title = f"{customer_name} Presentation {str(datetime.now())}"
    body = {
        'name': copy_title
    }
    drive_response = client.drive_service.files().copy(
        fileId=template_presentation_id, body=body).execute()
    presentation_copy_id = drive_response.get('id')

    try:
        requests = [
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{customer-name}',
                        'matchCase': True
                    },
                    'replaceText': customer_name
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{case-description}',
                        'matchCase': True
                    },
                    'replaceText': template_data["case_description"]
                }
            },
            {
                'replaceAllText': {
                    'containsText': {
                        'text': '{total-portfolio}',
                        'matchCase': True
                    },
                    'replaceText': template_data["total_portfolio"]
                }
            }
        ]

        update_presetation_with_requests(
            presentation_id=presentation_copy_id, client=client, requests=requests)

        print(f"Created presentation for "
              f"{customer_name} with ID: {presentation_copy_id}")

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


if __name__ == '__main__':
    client = GoogleClient(scopes=SCOPES)
    template_data = {
        "case_description": "This case is really good I think that it may be the best ive ever seen",
        "total_portfolio": "98%"
    }
    create_presentation_from_template(
        customer_name='Guild',
        template_presentation_id=PRESENTATION_ID,
        client=client,
        template_data=template_data)
