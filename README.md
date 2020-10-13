# Photoshop CAH Generator

![A quick peek of the script in action!](https://i.imgur.com/4CwHYiz.gif)

After struggling to come up with birthday gift ideas, I decided to make a custom set of Cards Against Humanity cards for my friends.

While there are generators online, I could not find one that would allow me to change the logo AND footer; they were usually restricted to one or the other. Thus, I decided to make my own using Photoshop since the card supplier provided a perfectly sized PSD template.

## How it works

I wanted this to be something that my friends could contribute ideas to, so I made a Google Sheet that collates everyones ideas and interfaced with it.

1. People populate the Google Sheet with the card text, card colour, who it's assigned to, and whether it's a special "Pick 2" or "Draw 2 Pick 3" card.

![The page where my friends can contribute to which generates the card data.](https://i.imgur.com/SMoU85N.png)

2. The cards to be generated are chosen with a checkbox. This is then used with a Query to filter out the cards that have been picked out.

    The formula: 
    ```
    =QUERY(ARRAYFORMULA(TO_TEXT({CardIdeaFormattedSpaces, CardIdeaColour, CardIdeaSpecial, 
    CardIdeaAssignee, 'Card ideas'!F2:F})), "Select Col1, Col2, Col3, Col4 WHERE Col5 = 'TRUE'")
    ```
3. The script goes to the `Cards to generate` sheet and reads it as JSON.
4. A JSON object is generated out of the Sheet, containing all the relevant information.
5. Using the Windows COM interface, the script opens the Photoshop template and iterates through the JSON object, hiding and showing layers as appropriate and saving each as an appropriately named PNG file.

## Acknowledgements

1. http://techarttiki.blogspot.com/2008/08/photoshop-scripting-with-python.html - For helping me get my head around COM interfaces.
2. The Adobe community forums.
3. The Adobe Developer Docs.