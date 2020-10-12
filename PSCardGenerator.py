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
    ignore_folders = ['P2B', 'D2P3B']
    layer_names = [layer.Name for layer in doc.Layers]
    all_layers = [
        layer for layer in layer_names if layer not in ignore_folders]
    # Initial hiding of all layers to handle colour switching
    for layer in all_layers:
        doc.ArtLayers[layer].visible = False
    layers_to_show = [
        layer for layer in all_layers if card_colour in layer if layer not in ignore_folders]
    special_layers = {'PICK 2': "BlackPick2Layer",
                      'DRAW 2 PICK 3': "BlackDraw3Layer"}
    for layer in layers_to_show:
        if card_assignee in layer:
            # Show appropriate face logo layer
            doc.ArtLayers[layer].visible = True
        # Show the rest of the layers
        if "CardFaceLogo" not in layer:
            doc.ArtLayers[layer].visible = True
    if card_special != "N/A":
        doc.ArtLayers[special_layers[card_special]].visible = True
    doc.ArtLayers[f"{card_colour}CardTextLayer"].TextItem.contents = card_text
    doc.ArtLayers[f"{card_colour}CardFooter"].TextItem.contents = card_footer
    doc.ArtLayers["HiddenLayerWhenFinished"].visible = False

    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13
    options.PNG8 = False
    pngfile = fr"E:\MyDocuments\Programming\MarkGithub\PhotoshopFiddlesGame\Card_output\{card_colour}_front_{index_number}.png"
    doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)
    print(f"{card_text} -- Exported PNG of {card_colour}_front_{index_number}_{card_assignee}")


print("Generating cards...")
card_json = json_generator()
for index, card in enumerate(card_json):
    template_fill(card['text'], card['colour'], card['special'],
                  card['assignee'], card['footer'], index)
print(f"========================\nComplete! {len(card_json)} cards generated!")
