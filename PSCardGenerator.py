import json
import win32com.client
import os
import requests
import itertools


def grouper(iterable, n, fillvalue=None):
    "Collect data into fixed-length chunks or blocks"
    # grouper('ABCDEFG', 3, 'x') --> ABC DEF Gxx"
    args = [iter(iterable)] * n
    return itertools.zip_longest(*args, fillvalue=fillvalue)


def json_generator():
    gsheets_url = 'https://spreadsheets.google.com/feeds/cells/1_O1vE52BJOpJcOs-kmh3N79q5PHFG9OhXJ0Yhm2KXq0/2/public/full?alt=json'
    r = requests.get(gsheets_url)
    data = r.json()
    card_entries = data['feed']['entry']
    card_data = []
    for x, y, z, a, b in grouper(card_entries[6:], 5):
        card_structure = {
            "text": x['gs$cell']['$t'],
            "colour": y['gs$cell']['$t'],
            "special": z['gs$cell']['$t'],
            "assignee": a['gs$cell']['$t'],
            "footer": b['gs$cell']['$t']
        }
        card_data.append(card_structure)
    return card_data


def template_fill(card_text, card_colour, card_special, card_assignee, card_footer, index_number):
    psApp = win32com.client.Dispatch("Photoshop.Application")
    psApp.Open(
        r"E:\MyDocuments\Programming\MarkGithub\PhotoshopFiddlesGame\Fiddles_CAH-Portrait.psd")
    doc = psApp.Application.ActiveDocument
    underscore_pattern = "_______"
    white_layers = ["WhiteCardFooter", "WhiteCardTextLayer",
                    "WhiteBG", f"WhiteCardFaceLogo{card_assignee}"]
    black_layers = ["BlackCardFooter", "BlackCardTextLayer",
                    "BlackBG", f"BlackCardFaceLogo{card_assignee}"]
    assignee_names = ['Alex', 'Fiddles', 'Ellie']
    # # looping through layer groups
    # if (len(doc.LayerSets) > 0):
    #     for group in doc.LayerSets:
    #         print(type(group.layers[1]))
    #         print(group.layers[1].name)
    # for layer in doc.Layers if 'White':
    #     print(layer.Name)
    # TODO: Look at using layer names to clean up relevant layer hiding code
    all_layers = [layer.Name for layer in doc.Layers]
    print([layer for layer in all_layers if "White" not in layer])
    special_layers = ["BlackPick2Layer", "BlackDraw3Layer"]
    for name in assignee_names:
        doc.ArtLayers[f"WhiteCardFaceLogo{name}"].visible = False
        doc.ArtLayers[f"BlackCardFaceLogo{name}"].visible = False
    for special_layer in special_layers:
        doc.ArtLayers[special_layer].visible = False
    if card_colour == "Black":
        print("Black card: Hiding white layers...")
        for white_layer in white_layers:
            doc.ArtLayers[white_layer].visible = False
        # set black card layers to be visible
        for black_layer in black_layers:
            doc.ArtLayers[black_layer].visible = True
        if card_text.count(underscore_pattern) < 2 or card_special == 'N/A':
            print("Using standard template...")
        if card_text.count(underscore_pattern) == 2 or card_special == 'PICK 2':
            doc.ArtLayers["BlackPick2Layer"].visible = True
            print("Using Pick Two template...")
        if card_text.count(underscore_pattern) == 3 or card_special == 'DRAW 2 PICK 3':
            doc.ArtLayers["BlackDraw3Layer"].visible = True
            print("Using Pick Three template...")
    elif card_colour == "White":
        print("White card: Hiding black layers")
        for black_layer in black_layers:
            doc.ArtLayers[black_layer].visible = False
        for white_layer in white_layers:
            doc.ArtLayers[white_layer].visible = True
    doc.ArtLayers[f"{card_colour}CardTextLayer"].TextItem.contents = card_text
    doc.ArtLayers[f"{card_colour}CardFooter"].TextItem.contents = card_footer
    doc.ArtLayers[f"WhiteCardFaceLogo{card_assignee}"].visible = True
    doc.ArtLayers["HiddenLayerWhenFinished"].visible = False

    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13
    options.PNG8 = False
    pngfile = fr"E:\MyDocuments\Programming\MarkGithub\PhotoshopFiddlesGame\Card_output\{card_colour}_front_{index_number}.png"
    doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)
    print(f"{card_text} -- Exported PNG of {card_colour}_front_{index_number}_{card_assignee}")


card_json = json_generator()
for index, card in enumerate(card_json):
    template_fill(card['text'], card['colour'], card['special'],
                  card['assignee'], card['footer'], index)
print(f"Complete! {len(card_json)} cards generated!")
