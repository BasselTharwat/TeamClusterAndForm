import pandas as pd
import random
import string

def generate_name():
    # Generate a random first and last name
    first_names = ['John', 'Jane', 'Michael', 'Emily', 'David', 'Sarah', 'James', 'Emma', 
                  'William', 'Olivia', 'Daniel', 'Sophia', 'Matthew', 'Isabella', 'Joseph',
                  'Mia', 'Andrew', 'Charlotte', 'Joshua', 'Amelia', 'Ryan', 'Harper', 'Nicholas',
                  'Evelyn', 'Tyler', 'Abigail', 'Alexander', 'Elizabeth', 'Nathan', 'Sofia']
    
    last_names = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis',
                 'Rodriguez', 'Martinez', 'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson',
                 'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin', 'Lee', 'Perez', 'Thompson',
                 'White', 'Harris', 'Sanchez', 'Clark', 'Ramirez', 'Lewis', 'Robinson']
    
    return f"{random.choice(first_names)} {random.choice(last_names)}"

def generate_dash_id():
    # Use one of 6 specific dash IDs
    dash_ids = ['52-', '61-', '64-', '58-', '55-']
    return random.choice(dash_ids)

def generate_gender():
    return random.choice(['male', 'female'])

def generate_type():
    return random.choice(['Leader', 'Subleader', 'None'])

def generate_american_phone_number():
    # American phone numbers: (XXX) XXX-XXXX
    area_code = random.randint(200, 999)
    prefix = random.randint(200, 999)
    line_number = random.randint(1000, 9999)
    return f"({area_code}) {prefix}-{line_number}"

def fuzzy_friend_name(name):
    # Randomly apply one of several fuzzy modifications
    mods = [
        lambda n: n[:max(1, len(n)-1)],  # drop last char
        lambda n: n[1:] if len(n) > 1 else n,  # drop first char
        lambda n: n.replace('a', '@', 1) if 'a' in n else n,  # replace a with @
        lambda n: n + random.choice(['-', '_', '']),  # add dash or underscore
        lambda n: n[:len(n)//2],  # use only first half
        lambda n: n[::-1],  # reverse
        lambda n: n,  # no change
    ]
    mod = random.choice(mods)
    return mod(name)

def generate_friends(all_names, self_name):
    # Generate 0-2 random friends from the list of names
    num_friends = random.randint(0, 2)
    if num_friends == 0:
        return ["", ""]
    
    available_friends = [name for name in all_names if name != self_name]  # Exclude self
    friends = random.sample(available_friends, min(num_friends, len(available_friends)))
    
    # Fuzzify the friend names
    fuzzy_friends = [fuzzy_friend_name(friend) for friend in friends]
    
    # Pad with empty strings if needed
    while len(fuzzy_friends) < 2:
        fuzzy_friends.append("")
    
    return fuzzy_friends

# Generate 100 names first
names = [generate_name() for _ in range(100)]

# Create the data
data = []
for name in names:
    friends = generate_friends(names, name)
    data.append({
        'Name': name,
        'Dash ID': generate_dash_id(),
        'Gender': generate_gender(),
        'Type': generate_type(),
        'Number': generate_american_phone_number(),
        'Friend 1': friends[0],
        'Friend 2': friends[1],
    })

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
df.to_excel('test_data_new.xlsx', index=False)
print("Test data generated and saved to 'test_data_new.xlsx'") 