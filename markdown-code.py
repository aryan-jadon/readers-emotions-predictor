from __future__ import print_function, division, unicode_literals
from module_files import example_helper
import json
from nltk.tokenize import sent_tokenize
from tqdm import tqdm
import numpy as np
import emoji
import os
import operator
from torchmoji.sentence_tokenizer import SentenceTokenizer
from torchmoji.model_def import torchmoji_emojis
from torchmoji.global_variables import PRETRAINED_PATH, VOCAB_PATH
import docx
from docx import Document
from docx.shared import Inches
import warnings
from mdutils.mdutils import MdUtils
from mdutils import Html

warnings.filterwarnings('ignore')

EMOJIS = ":joy: :unamused: :weary: :sob: :heart_eyes: \
:pensive: :ok_hand: :blush: :heart: :smirk: \
:grin: :notes: :flushed: :100: :sleeping: \
:relieved: :relaxed: :raised_hands: :two_hearts: :expressionless: \
:sweat_smile: :pray: :confused: :kissing_heart: :heartbeat: \
:neutral_face: :information_desk_person: :disappointed: :see_no_evil: :tired_face: \
:v: :sunglasses: :rage: :thumbsup: :cry: \
:sleepy: :yum: :triumph: :hand: :mask: \
:clap: :eyes: :gun: :persevere: :smiling_imp: \
:sweat: :broken_heart: :yellow_heart: :musical_note: :speak_no_evil: \
:wink: :skull: :confounded: :smile: :stuck_out_tongue_winking_eye: \
:angry: :no_good: :muscle: :facepunch: :purple_heart: \
:sparkling_heart: :blue_heart: :grimacing: :sparkles:".split(' ')


def top_elements(array, k):
    ind = np.argpartition(array, -k)[-k:]
    return ind[np.argsort(array[ind])][::-1]


with open(VOCAB_PATH, 'r') as f:
    vocabulary = json.load(f)


def get_reactions(vocabulary, predict_text, reactionTracker):
    maxlen = len(predict_text)
    st = SentenceTokenizer(vocabulary, maxlen)

    # Loading model
    model = torchmoji_emojis(PRETRAINED_PATH)
    # Running predictions
    tokenized, _, _ = st.tokenize_sentences([predict_text])

    # Get sentence probability
    prob = model(tokenized)[0]

    # Top emoji id
    emoji_ids = top_elements(prob, 5)

    # map to emojis
    emojis = map(lambda x: EMOJIS[x], emoji_ids)

    Predicted = emoji.emojize("{}".format(' '.join(emojis)), use_aliases=True)

    for react in Predicted.split(' '):
        if react in reactionTracker:
            reactionTracker[react] += 1
        else:
            reactionTracker[react] = 1

    return predict_text, Predicted


def make_report(book_path):
    doc = docx.Document(book_path)
    reactionTracker = {}
    name = book_path.split('/')[-1].split('.docx')[0]
    temp_title_path = name

    novel_title = temp_title_path
    novel_title = novel_title.replace('-', ' ')
    complete_novel_title = "Analysis of Book - " + str(novel_title.capitalize())

    mdFile = MdUtils(file_name=str("results/")+str(temp_title_path), title=complete_novel_title)

    # Adding Image
    path = "../images/image-4.jpg"
    mdFile.new_paragraph(Html.image(path=path, size='800x800', align='center'))
    mdFile.new_paragraph("<div style='page-break-after: always;'> </div>")

    for para in doc.paragraphs:
        if len(para.text) != 0:
            checktext = para.text.strip()
            checktext = checktext.replace(" ", "")

            if not checktext.isnumeric():
                try:
                    tempText = para.text
                    tempText += os.linesep
                    predictions = get_reactions(
                        vocabulary, para.text, reactionTracker)
                    tempText += " ".join(predictions[1].split(','))

                    mdFile.new_paragraph(tempText)
                    mdFile.new_paragraph()
                except:
                    tempText = para.text
                    tempText += os.linesep
                    mdFile.new_paragraph(tempText)

    mdFile.new_paragraph("<div style='page-break-after: always;'> </div>")
    mdFile.new_header(3, "Final Count of all the reactions on context by our engine.")

    path = "../Images/Image-10.jpg"
    mdFile.new_paragraph(Html.image(path=path, size='800x800', align='center'))
    mdFile.new_paragraph("<div style='page-break-after: always;'> </div>")

    mdFile.new_header(3, "Reactions Count.")

    from collections import OrderedDict

    reactions_sorted_by_value = OrderedDict(sorted(reactionTracker.items(),
                                                   key=lambda x: x[1], reverse=True))
    reactions_icon = []
    reactions_count = []

    for icon, value in reactions_sorted_by_value.items():
        reactions_icon.append(icon)
        reactions_count.append(value)

    list_of_strings = ["Reactions", "Count"]

    for row in range(len(reactions_icon)):
        list_of_strings.extend([str(reactions_icon[row]), str(reactions_count[row])])

    mdFile.new_table(columns=2, rows=len(reactions_count)+1, text=list_of_strings, text_align='center')
    mdFile.create_md_file()


booksList = os.listdir('results/')

for book_path in tqdm(range(len(booksList))):
    path = "Completed/"+str(booksList[book_path])
    try:
        print("Working on " + str(booksList[book_path]))
        make_report(path)
    except Exception as e:
        print(e)
