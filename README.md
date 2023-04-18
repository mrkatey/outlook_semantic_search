# outlook_semantic_search
Find your messages, using Semantic search not just keywords! :)

# A step-by-step guide

## Extract your outlook messages

1. Open Microsoft Outlook on your computer.

2. Click on the "File" menu in the top left corner of the window.

3. Click on "Open & Export" in the left-hand menu.

4. Click on "Import/Export" in the right-hand menu.

5. In the "Import and Export Wizard" window that opens, select "Export to a file" and click "Next."

6. Select "Outlook Data File (.pst)" and click "Next."

7. Select the email account that you want to export and make sure the "Include subfolders" option is checked. Click "Next."

8. Choose a location and file name to save the exported file, and click "Finish."

9. If you want to password-protect your exported file, enter and confirm a password, and click "OK."

10. Outlook will begin exporting your messages to the specified location. Depending on the size of your mailbox, this may take some time.

Once the export is complete, you will have a .pst file that contains all of your Outlook messages, including emails, contacts, calendar items, and tasks. 

## Create a model of your messages

Run the following Python script from your commandline:
```python
import win32com.client
import os
import gensim
from gensim import corpora, models, similarities

# Path to the folder containing the Outlook PST file
folder_path = "C:/Users/username/Documents/Outlook Files/"

# Path to the output file containing the model
model_path = "C:/Users/username/Documents/outlook_model"

# Load the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Load the default email account
inbox = outlook.GetDefaultFolder(6)

# Get a list of all email messages in the inbox
messages = inbox.Items

# Create a list to store the text of each message
texts = []

# Loop through all messages and extract the body text
for message in messages:
    try:
        body = message.Body
        texts.append(body)
    except:
        pass

# Preprocess the text data
texts = [gensim.utils.simple_preprocess(text) for text in texts]

# Create a dictionary and corpus of the preprocessed texts
dictionary = corpora.Dictionary(texts)
corpus = [dictionary.doc2bow(text) for text in texts]

# Build a TF-IDF model of the corpus
tfidf = models.TfidfModel(corpus)

# Build an LSI model of the TF-IDF corpus
lsi = models.LsiModel(tfidf[corpus], id2word=dictionary, num_topics=200)

# Save the LSI model to disk
lsi.save(model_path)

# Load the LSI model from disk
lsi = models.LsiModel.load(model_path)

# Build an index of the LSI model
index = similarities.MatrixSimilarity(lsi[corpus])

# Define a function to perform a semantic search
def semantic_search(query):
    query = gensim.utils.simple_preprocess(query)
    query_bow = dictionary.doc2bow(query)
    query_lsi = lsi[query_bow]
    sims = index[query_lsi]
    sims = sorted(enumerate(sims), key=lambda item: -item[1])
    results = []
    for sim in sims:
        if sim[1] > 0.2:
            results.append((sim[1], texts[sim[0]]))
    return results

# Perform a semantic search for a query
query = "important project deadline"
results = semantic_search(query)
for result in results:
    print(result)

```
