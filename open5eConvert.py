import pandas as pd
import requests
import sys
import getopt
from docx import Document

def speed_to_text(speed):
    for k, v in speed.items():
        if k not in ['walk', 'burrow', 'climb', 'fly', 'swim', 'hover']:
            sys.exit(f'invalid speed key found: {k}')
    speed_text = ""
    if 'walk' in speed:
        speed_text = speed_text + f"{speed['walk']} ft."
    if 'burrow' in speed:
        if speed_text != "":
            speed_text = speed_text + ", "
        speed_text = speed_text + f"Burrow {speed['burrow']} ft."
    if 'climb' in speed:
        if speed_text != "":
            speed_text = speed_text + ", "
        speed_text = speed_text + f"Climb {speed['climb']} ft."
    if 'fly' in speed:
        if speed_text != "":
            speed_text = speed_text + ", "
        speed_text = speed_text + f"Fly {speed['fly']} ft.{' (hover)' if 'hover' in speed and speed['hover'] else ''}"
    if 'swim' in speed:
        if speed_text != "":
            speed_text = speed_text + ", "
        speed_text = speed_text + f"Swim {speed['swim']} ft."
    return speed_text

def attr_to_text(attr):
    modifier = int(attr/2) - 5
    return f"{attr} ({modifier:+})"

def save_to_text(monster):
    save_text = ""
    if monster['strength_save']:
        save_text = save_text + f"Str {monster['strength_save']:+}"
    if monster['dexterity_save']:
        if save_text != "":
            save_text = save_text + ", "
        save_text = save_text + f"Dex {monster['dexterity_save']:+}"
    if monster['constitution_save']:
        if save_text != "":
            save_text = save_text + ", "
        save_text = save_text + f"Con {monster['constitution_save']:+}"
    if monster['intelligence_save']:
        if save_text != "":
            save_text = save_text + ", "
        save_text = save_text + f"Int {monster['intelligence_save']:+}"
    if monster['wisdom_save']:
        if save_text != "":
            save_text = save_text + ", "
        save_text = save_text + f"Wis {monster['wisdom_save']:+}"
    if monster['charisma_save']:
        if save_text != "":
            save_text = save_text + ", "
        save_text = save_text + f"Cha {monster['charisma_save']:+}"
    return save_text

def setup_document(title=""):
    # Create a new Word document
    doc = Document()
    doc.add_heading(title, level=0)
    return(doc)

def monster_import(results, doc):
    doc.add_heading("Monsters", level=1)
    for monster in results:
        print('Importing ' + str(monster['name']))
        doc.add_heading(str(monster['name']), level=2)
        doc.add_paragraph(f"{monster['size'] if monster['size'].lower() != 'titanic' else 'Gargantuan'} {monster['type']}, {monster['alignment']}")
        doc.add_paragraph(f"Armor Class {monster['armor_class']}{monster['armor_desc'] if (monster['armor_desc']) else ''}")
        doc.add_paragraph(f"Hit Points {monster['hit_points']} ({monster['hit_dice']})")
        doc.add_paragraph("Speed " + speed_to_text(monster['speed']))
        doc.add_paragraph("STR DEX CON INT WIS CHA")
        doc.add_paragraph(attr_to_text(monster['strength']) + " " +
                attr_to_text(monster['dexterity']) + " " +
                attr_to_text(monster['constitution']) + " " +
                attr_to_text(monster['intelligence']) + " " +
                attr_to_text(monster['wisdom']) + " " +
                attr_to_text(monster['charisma']) )
        save_text = save_to_text(monster)
        if save_text != "":
            doc.add_paragraph("Saving Throws " + save_to_text(monster))
        # doc.add_paragraph("Skills") - they don't seem to be used
        doc.add_paragraph("Damage Vulnerabilities  " + str(monster['damage_vulnerabilities']))
        doc.add_paragraph("Damage Resistances  " + str(monster['damage_resistances']))
        doc.add_paragraph("Damage Immunities " + str(monster['damage_immunities']))
        doc.add_paragraph("Condition Immunities " + str(monster['condition_immunities']))
        doc.add_paragraph("Senses " + str(monster['senses']))
        doc.add_paragraph("Languages " + str(monster['languages']))
        doc.add_paragraph("Challenge " + str(monster['challenge_rating']))

        if monster['actions']:
            doc.add_heading("Actions", level=3)
            for index, a in enumerate(monster['actions']):
                p = doc.add_paragraph("")
                runner = p.add_run(a['name'] + ". ")
                runner.bold = True
                runner.italic = True
                p.add_run(a['desc'])

        if monster['bonus_actions']:
            doc.add_heading("Bonus Actions", level=3)
            for index, a in enumerate(monster['bonus_actions']):
                p = doc.add_paragraph("")
                runner = p.add_run(a['name'] + ". ")
                runner.bold = True
                runner.italic = True
                p.add_run(a['desc'])

        if monster['reactions']:
            doc.add_heading("Reactions", level=3)
            for index, a in enumerate(monster['reactions']):
                p = doc.add_paragraph("")
                runner = p.add_run(a['name'] + ". ")
                runner.bold = True
                runner.italic = True
                p.add_run(a['desc'])

        if monster['legendary_actions']:
            doc.add_heading("Legendary Actions", level=3)
            if monster['legendary_desc']:
                doc.add_paragraph(monster['legendary_desc'])
            for index, a in enumerate(monster['legendary_actions']):
                p = doc.add_paragraph("")
                runner = p.add_run(a['name'] + ". ")
                runner.bold = True
                runner.italic = True
                p.add_run(a['desc'])

        if monster['special_abilities']:
            doc.add_heading("Special Abilities", level=3)
            for index, a in enumerate(monster['special_abilities']):
                p = doc.add_paragraph("")
                runner = p.add_run(a['name'] + ". ")
                runner.bold = True
                runner.italic = True
                p.add_run(a['desc'])

def spell_import(results, doc):
    doc.add_heading("Spells", level=1)
    for spell in results:
        print('Importing ' + str(spell['name']))
        doc.add_heading(str(spell['name']), level=2)
        if (spell['level'] == 'Cantrip'):
            doc.add_paragraph(f"{spell['school']} {spell['level']}")
        else:
            doc.add_paragraph(f"{spell['level']} {spell['school']} {'(ritual}' if spell['ritual'] == 'yes' else ''}")
        doc.add_paragraph(f"Casting Time {spell['casting_time']}")
        doc.add_paragraph(f"Range {spell['range']}")
        doc.add_paragraph(f"Components {spell['components']}{' (' + spell['material'] + ')' if spell['material'] != '' else ''}")
        doc.add_paragraph(f"Duration {'concentration, ' if spell['concentration'] == 'yes' else '' }{spell['duration']}")
        doc.add_paragraph(f"{'classes ' + spell['dnd_class'] if spell['dnd_class'] != '' else ''}")
        doc.add_paragraph(spell['desc'])
        doc.add_paragraph(spell['higher_level'])

    doc.add_heading("Artificer", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Artificer') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Wizard", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Wizard') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Sorcerer", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Sorceror') >= 0 or spell['dnd_class'].find('Sorcerer') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Cleric", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Cleric') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Ranger", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Ranger') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Warlock", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Warlock') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Bard", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Bard') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Paladin", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Paladin') >= 0 or spell['dnd_class'].find('Herald') >= 0):
            doc.add_paragraph(spell['name'])
    doc.add_heading("Druid", level=1)
    for spell in results:
        if (spell['dnd_class'].find('Druid') >= 0):
            doc.add_paragraph(spell['name'])


def item_import(results, doc):
    doc.add_heading("Items", level=1)
    for item in results:
        print('Importing ' + str(item['name']))
        doc.add_heading(str(item['name']), level=2)
        doc.add_paragraph(f"{item['type']}, {item['rarity']} {'(' + item['requires_attunment'] + ')' if item['requires_attunment'] != '' else ''}")
        doc.add_paragraph(item['desc'])

def feat_import(results, doc):
    doc.add_heading("Feats", level=1)
    for feat in results:
        print('Importing ' + str(feat['name']))
        doc.add_heading(str(feat['name']), level=2)
        if (feat['prerequisite'] != ''):
            doc.add_paragraph(f"Prerequisite {feat['prerequisite']}")
        doc.add_paragraph(feat['desc'])

def main():
    argv = sys.argv[1:]

    opts, args = getopt.getopt(
        argv,
        "t:f:s:o:h",
        ["title=", "feature=", "source=", "out=", "help"]
    )

    feature = ""
    source = ""
    doc_title = ""
    doc_filename = ""
    for opt, arg in opts:
        if opt in ['-f', '--feature']:
            feature = arg.lower()
            if feature not in ['monsters', 'spells', 'magicitems', 'feats']:
                sys.exit(f"Invalid feature {feature}. Valid features are " +
                         "monsters, spells, items, and feats.")
            print(feature)
        if opt in ['-s', '--source']:
            source = arg.lower()
            print(source)
        if opt in ['-t', '--title']:
            doc_title = arg
            print(doc_title)
        if opt in ['-o', '--out']:
            doc_filename = arg
            print(doc_filename)
    source_name = ""
    if (source == 'menagerie'):
        source_name = 'Level Up Advanced 5e Monstrous Menagerie'
    elif (source == 'wotc-srd'):
        source_name = '5e Core Rules'
    elif (source == 'a5e'):
        source_name = 'Level Up Advanced 5e'
    elif (source == 'dmag'):
        source_name = 'Deep Magic 5e'
    elif (source == 'dmag-e'):
        source_name = 'Deep Magic Extended'
    elif (source == 'warlock'):
        source_name = 'Warlock Archives'
    elif (source == 'kp'):
        source_name = 'Kobold Press Compilation'
    elif (source == 'toh'):
        source_name = 'Tome of Heroes'
    elif (source == 'tob'):
        source_name = 'Tome of Beasts'
    elif (source == 'tob2'):
        source_name = 'Tome of Beasts 2'
    elif (source == 'tob3'):
        source_name = 'Tome of Beasts 3'
    elif (source == 'cc'):
        source_name = 'Creature Codex'
    elif (source == 'taldorei'):
        source_name = 'Critical Role: Tal’Dorei Campaign Setting'
    elif (source == 'vom'):
        source_name = 'Vault of Magic'
    else:
        sys.exit(f"Invalid source {source}. Valid sources are " +
                    "menagerie, wotc-srd, a5e, dmag, dmag-e, warlock, kp, toh, tob, tob2, tob3, cc, taldorei, vom.")
    if (doc_title == ""):
        doc_title = f"{source.capitalize() if source_name == '' else source_name} {feature.capitalize()}"
        print(f"No title provided, using default title '{doc_title}'")
    if (doc_filename == ""):
        doc_filename = f"{doc_title}.docx"
        print(f"No filename provided, using default filename '{doc_filename}'")

    doc = setup_document(title=doc_title)

    # load data from API
    page_no = 0
    params = {'document__slug__iexact': source, 'page': page_no}
    results = []
    sys.stdout.write(f"Downloading https://api.open5e.com/v1/{feature}")
    while True:
        params['page'] += 1
        try:
            sys.stdout.write(".")
            sys.stdout.flush()
            r = requests.get(f"https://api.open5e.com/v1/{feature}", params=params)
            r.raise_for_status()
        except requests.exceptions.HTTPError as err:
            break

        df = pd.DataFrame(r.json())
        for x in df['results'].items():
            results.append(x[1])
    sys.stdout.write("Done!\n")

    if (feature == 'monsters'):
        print("Importing Monsters")
        monster_import(results, doc)
    elif (feature == 'spells'):
        print("Importing Spells")
        spell_import(results, doc)
    elif (feature == 'magicitems'):
        print("Importing Magic Items")
        item_import(results, doc)
    elif (feature == 'feats'):
        print("Importing feats")
        feat_import(results, doc)


    # Save the Word document
    doc.save(doc_filename)
    print ("wrote file " + doc_filename)


main()
