import pandas as pd
import random
from collections import defaultdict
import os

# === Set your input file path here ===
input_path = r'C:\Users\lsandow\Documents\GitHub\Full-Magic-Pack\Sets\DM\Python\duel_masters.xlsx'

# === Generate a unique output filename ===
base_output = os.path.join(os.path.dirname(input_path), 'generated_packs.xlsx')
output_path = base_output
counter = 1
while os.path.exists(output_path):
    output_path = base_output.replace('.xlsx', f'{counter}.xlsx')
    counter += 1

# Load data
df = pd.read_excel(input_path)
df.columns = [col.strip().title() for col in df.columns]

# Parse color list
df['Color'] = df['Color'].fillna('').apply(lambda x: [c.strip() for c in x.split(',') if c.strip()])
df['Has_Land_Back'] = df.get('Has_Land_Back', False).astype(bool)
df['Is_Dual_Land'] = df.get('Is_Dual_Land', False).astype(bool)

# Filter groups by rarity (excluding dual lands from common slot)
rarity_groups = {
    'Common': df[(df['Rarity'].str.lower() == 'common') & (~df['Is_Dual_Land'])],
    'Uncommon': df[df['Rarity'].str.lower() == 'uncommon'],
    'Rare': df[df['Rarity'].str.lower() == 'rare'],
    'Mythic': df[df['Rarity'].str.lower() == 'mythic'],
    'Special': df[df['Rarity'].str.lower() == 'special']
}

dual_lands = df[df['Is_Dual_Land']]
land_back = df[df['Has_Land_Back']]

usage_limits = {'Common': 5, 'Uncommon': 3, 'Rare': 2, 'Mythic': 1, 'Special': 1}
usage_count = defaultdict(int)
color_count = defaultdict(int)

def choose_rarity(rarity_probs):
    r = random.random()
    cumulative = 0
    # Shuffle order of rarity checking
    for rarity, prob in random.sample(list(rarity_probs.items()), len(rarity_probs)):
        cumulative += prob
        if r <= cumulative:
            return rarity
    return "Common"  # fallback

def get_random_card(pool, rarity, used_cards):
    options = pool[
        (~pool['Name'].isin(used_cards)) &
        (pool['Name'].apply(lambda name: usage_count[name] < usage_limits[rarity]))
    ]
    if options.empty:
        return None
    return options.sample(frac=1).sample(1).iloc[0]

def update_color_count(card):
    for color in card['Color']:
        color_count[color] += 1

def generate_single_pack():
    pack = []
    used_this_pack = set()
    colors_needed = ['White', 'Blue', 'Black', 'Red', 'Green']

    def add_card(card, slot=None):
        used_this_pack.add(card['Name'])
        usage_count[card['Name']] += 1
        update_color_count(card)
        card_copy = card.copy()
        card_copy['Slot'] = slot
        pack.append(card_copy)

    def pick_one_mono_per_color(rarity, total_count):
        pool = rarity_groups[rarity]
        for color in random.sample(['White', 'Blue', 'Black', 'Red', 'Green'], 5):
            if len([c for c in pack if c['Rarity'].lower() == rarity.lower()]) >= total_count:
                break
            candidates = pool[
                (pool['Color'].apply(lambda c: len(c) == 1 and c[0] == color)) &
                (~pool['Name'].isin(used_this_pack)) &
                (pool['Name'].apply(lambda name: usage_count[name] < usage_limits[rarity]))
            ]
            if not candidates.empty:
                card = candidates.sample(1).iloc[0]
                add_card(card, slot=rarity)

        while len([c for c in pack if c['Rarity'].lower() == rarity.lower()]) < total_count:
            candidates = pool[
                (~pool['Name'].isin(used_this_pack)) &
                (pool['Name'].apply(lambda name: usage_count[name] < usage_limits[rarity]))
            ]
            if candidates.empty:
                break
            card = candidates.sample(1).iloc[0]
            add_card(card, slot=rarity)

    # Commons and uncommons with color balance
    pick_one_mono_per_color('Common', 7)
    pick_one_mono_per_color('Uncommon', 3)

    # 8th common slot: 12% chance to pull from Special rarity
    if random.random() <= 0.12:
        # Use a Special rarity card
        special_pool = rarity_groups['Special']
        card = get_random_card(special_pool, 'Special', used_this_pack)
        if card is not None and not card.empty:
            add_card(card, slot='Special')
        else:
            # fallback to normal common if no specials available
            card = get_random_card(rarity_groups['Common'], 'Common', used_this_pack)
            if card is not None and not card.empty:
                add_card(card, slot='Common 8')
    else:
        card = get_random_card(rarity_groups['Common'], 'Common', used_this_pack)
        if card is not None and not card.empty:
            add_card(card, slot='Common 8')

    # Rare or mythic
    if random.randint(1, 6) == 1:
        rarity = 'Mythic'
    else:
        rarity = 'Rare'
    card = get_random_card(rarity_groups[rarity], rarity, used_this_pack)
    if card is not None and not card.empty:
        add_card(card, slot=rarity)

    # Wildcard 1
    wc1 = choose_rarity({'Common': 0.20, 'Uncommon': 0.60, 'Rare': 0.18, 'Mythic': 0.02})
    card = get_random_card(rarity_groups[wc1], wc1, used_this_pack)
    if card is not None and not card.empty:
        add_card(card, slot='Wildcard 1')

    # Wildcard 2
    wc2 = choose_rarity({'Common': 0.56, 'Uncommon': 0.36, 'Rare': 0.07, 'Mythic': 0.01})
    card = get_random_card(rarity_groups[wc2], wc2, used_this_pack)
    if card is not None and not card.empty:
        add_card(card, slot='Wildcard 2')

    # Land slot
    land_pool = dual_lands if random.random() <= 0.40 else land_back
    land_candidates = land_pool[~land_pool['Name'].isin(used_this_pack)]
    if not land_candidates.empty:
        card = land_candidates.sample(1).iloc[0]
        add_card(card, slot='Land')

    return pack

def check_color_balance():
    total_colored = sum(color_count.values())
    if total_colored == 0:
        return pd.DataFrame()
    avg = total_colored / len(color_count)
    balance_data = []

    for color, count in color_count.items():
        pct = (count / total_colored) * 100
        deviation = (count - avg) / avg * 100
        balance_data.append({
            'Color': color,
            'Used': count,
            'Share %': round(pct, 1),
            'Deviation %': round(deviation, 1),
            '⚠️': 'Yes' if abs(deviation) > 20 else ''
        })

    return pd.DataFrame(balance_data).sort_values(by='Deviation %', ascending=False)

def main():
    try:
        num_packs = int(input("How many booster packs would you like to generate? "))
    except ValueError:
        print("Please enter a valid integer.")
        return

    all_packs = []
    pack_rows = []

    for i in range(num_packs):
        pack = generate_single_pack()
        if pack is None or len(pack) < 13:
            print(f"⚠️ Warning: Pack {i+1} incomplete or not generated.")
        all_packs.append(pack)
        for card in pack:
            pack_rows.append({
                'Pack': f'Pack {i+1}',
                'Slot': card.get('Slot', ''),
                'Name': card['Name'],
                'Rarity': card['Rarity'],
                'Mana Cost': card['Mana Cost'],
                'Card Type': card['Card Type']
            })

    usage_summary = pd.DataFrame([
        {
            'Name': name,
            'Times Used': count,
            'Rarity': df[df['Name'] == name]['Rarity'].values[0] if not df[df['Name'] == name].empty else "Unknown"
        }
        for name, count in usage_count.items() if count > 0
    ]).sort_values(by='Times Used', ascending=False)

    print("\n📊 High-Usage Cards:")
    for _, row in usage_summary.iterrows():
        r = row['Rarity']
        cap = usage_limits.get(r, 99)
        warning = "⚠️" if row['Times Used'] >= cap else ""
        print(f"{row['Name']}: {row['Times Used']} used ({r}) {warning}")

    color_df = check_color_balance()
    if not color_df.empty:
        print("\n🎨 Color Balance:")
        print(color_df.to_string(index=False))

    with pd.ExcelWriter(output_path) as writer:
        pd.DataFrame(pack_rows).to_excel(writer, sheet_name='Booster Packs', index=False)
        usage_summary.to_excel(writer, sheet_name='Card Usage Summary', index=False)
        if not color_df.empty:
            color_df.to_excel(writer, sheet_name='Color Balance', index=False)

    print(f"\n✅ Saved to '{output_path}'")

if __name__ == '__main__':
    main()
