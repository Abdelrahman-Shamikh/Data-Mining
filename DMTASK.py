import pandas as pd

def eclat_from_excel(file_path, min_support, min_confidence):
    # Read data from Excel sheet
    df = pd.read_excel(file_path)

    # Assuming the columns are named 'TiD' and 'items'
    # Extract items from the "items" column
    dataset = [list(filter(str.isalpha, str(items))) for items in df['items']]

    eclat_from_dataset(dataset, min_support, min_confidence)

def eclat_from_dataset(dataset, min_support, min_confidence):
    # Call the eclat function
    freq_items_eclat = eclat(dataset, min_support)

    # Print frequent itemsets from Eclat
    print_freq_items_eclat(freq_items_eclat)

    # Generate and print strong association rules
    strong_rules = generate_strong_rules_with_confidence(dataset, freq_items_eclat, min_confidence)
    print_strong_rules(strong_rules)

    # Represent frequent items as association rules
    association_rules_representation = represent_as_association_rules(freq_items_eclat)
    print_association_rules_representation(association_rules_representation)

def eclat(dataset, min_support):
    items = {}
    transactions = len(dataset)

    for i, trans in enumerate(dataset, start=1):
        for item in trans:
            if frozenset({item}) not in items:
                items[frozenset({item})] = set()
            items[frozenset({item})].add(i)

    freq_items = {}
    l = 1
    while l == 1:
        l = 0
        new_items = {}
        for item1 in items:
            if len(items[item1]) >= min_support:
                freq_items[item1] = len(items[item1])

            for item2 in items:
                if item1 != item2:
                    union_set = item1.union(item2)
                    if union_set not in new_items:
                        new_items[union_set] = set(items[item1]).intersection(items[item2])
                        l = 1
        items = new_items

    return freq_items

def generate_strong_rules_with_confidence(dataset, frequent_itemsets, min_confidence):
    strong_rules = {}

    for itemset in frequent_itemsets:
        if len(itemset) > 1:
            for item in itemset:
                antecedent = frozenset(set(itemset) - {item})
                consequent = frozenset({item})
                confidence = calculate_confidence(dataset, antecedent, consequent)
                support = frequent_itemsets[itemset]

                if confidence >= min_confidence:
                    strong_rules[antecedent, consequent] = (confidence, support)

    return strong_rules

def calculate_confidence(dataset, antecedent, consequent):
    antecedent_support = len([1 for transaction in dataset if antecedent.issubset(transaction)])
    consequent_support = len([1 for transaction in dataset if consequent.issubset(transaction)])
    
    return consequent_support / antecedent_support

def represent_as_association_rules(freq_items):
    association_rules_representation = {}

    for itemset, support in freq_items.items():
        if len(itemset) > 1:
            for item in itemset:
                antecedent = frozenset(set(itemset) - {item})
                consequent = frozenset({item})
                association_rules_representation[antecedent, consequent] = support

    return association_rules_representation

def print_freq_items_eclat(freq_items_eclat):
    print("*************print_freq_items_eclat******************")
    for item, support in freq_items_eclat.items():
        print(f"{item}: {support}")

def print_strong_rules(rules):
    print("*****************************************************")
    print("*************print_strong_rules******************")
    for (antecedent, consequent), (confidence, support) in rules.items():
        print(f"Antecedent: {antecedent}, Consequent: {consequent}, Confidence: {confidence}, Support: {support}")

def print_association_rules_representation(rules_representation):
    print("*****************************************************")
    print("*************print_association_rules_representation******************")
    for (antecedent, consequent), support in rules_representation.items():
        print(f"Antecedent: {antecedent}, Consequent: {consequent}, Support: {support}")

def main():
    # Example usage with an absolute path using double backslashes
    file_path = r'D:\DMTASK\Horizontal_Format.xlsx'
    print("****************************************************************")
    min_support = int(input("Enter a min_support: "))
    min_confidence = int(input("Enter a min_confidence: "))
    print("****************************************************************")
    eclat_from_excel(file_path, min_support, min_confidence)
    print("****************************************************************")

if __name__ == "__main__":
    main()
