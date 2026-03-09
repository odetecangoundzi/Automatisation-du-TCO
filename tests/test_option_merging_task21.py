from decimal import Decimal

import pandas as pd

from core.merger import merge_company_into_tco


def test_option_grouping():
    # 1. Créer un DataFrame TCO de base minimal
    tco_data = [
        {
            "Code": "01",
            "Désignation": "LOT 01",
            "row_type": "section_header",
            "parent_code": "",
            "is_option": False,
        },
        {
            "Code": "01.1",
            "Désignation": "Article 1",
            "row_type": "article",
            "parent_code": "01",
            "is_option": False,
        },
        {
            "Code": "01",
            "Désignation": "Total LOT 01",
            "row_type": "recap",
            "parent_code": "01",
            "is_option": False,
        },
    ]
    merged_df = pd.DataFrame(tco_data)

    # 2. DPGF Compagnie 1 avec une option hors-bordereau + un article standard pour le match rate
    dpgf1_data = [
        {
            "Code": "01.1",
            "Désignation": "Article 1",
            "Qu.": Decimal("10"),
            "U": "m2",
            "Px_U_HT": Decimal("100"),
            "Px_Tot_HT": Decimal("1000"),
            "row_type": "article",
            "is_option": False,
        },
        {
            "Code": "OPT1",
            "Désignation": "Option Peinture Salon",
            "Qu.": Decimal("10"),
            "U": "m2",
            "Px_U_HT": Decimal("20"),
            "Px_Tot_HT": Decimal("200"),
            "row_type": "article",
            "is_option": True,
        },
    ]
    dpgf1 = pd.DataFrame(dpgf1_data)

    # 3. DPGF Compagnie 2 avec la MÊME option + le même article standard
    dpgf2_data = [
        {
            "Code": "01.1",
            "Désignation": "Article 1",
            "Qu.": Decimal("10"),
            "U": "m2",
            "Px_U_HT": Decimal("110"),
            "Px_Tot_HT": Decimal("1100"),
            "row_type": "article",
            "is_option": False,
        },
        {
            "Code": "VAR_P",
            "Désignation": "Option Peinture Salon",
            "Qu.": Decimal("10"),
            "U": "m2",
            "Px_U_HT": Decimal("25"),
            "Px_Tot_HT": Decimal("250"),
            "row_type": "article",
            "is_option": True,
        },
    ]
    dpgf2 = pd.DataFrame(dpgf2_data)

    # 4. Fusion Compagnie 1
    merged_df, alerts1 = merge_company_into_tco(merged_df, dpgf1, "Compagnie A", 0.20)
    print("Après Compagnie A, lignes:", len(merged_df))
    # On s'attend à ce que OPT_DYN soit créé
    opt_rows = merged_df[merged_df["parent_code"] == "OPT_DYN"]
    print("Nombre d'options:", len(opt_rows))

    # 5. Fusion Compagnie 2
    merged_df, alerts2 = merge_company_into_tco(merged_df, dpgf2, "Compagnie B", 0.20)
    print("Après Compagnie B, lignes:", len(merged_df))

    # 6. Vérification : On doit toujours avoir UNE SEULE ligne d'article dans OPT_DYN
    opt_articles = merged_df[
        (merged_df["parent_code"] == "OPT_DYN") & (merged_df["row_type"] == "article")
    ]
    print("Nombre d'articles d'options final:", len(opt_articles))

    if len(opt_articles) == 1:
        print("SUCCÈS : Regroupement par désignation fonctionnel !")
        row = opt_articles.iloc[0]
        print(f"Prix A: {row['Compagnie A_Px_U_HT']}, Prix B: {row['Compagnie B_Px_U_HT']}")
        assert row["Compagnie A_Px_U_HT"] == Decimal("20")
        assert row["Compagnie B_Px_U_HT"] == Decimal("25")
    else:
        print("ÉCHEC : Les options n'ont pas été regroupées.")
        for _i, r in opt_articles.iterrows():
            print(f"- {r['Code']}: {r['Désignation']}")


if __name__ == "__main__":
    try:
        test_option_grouping()
    except Exception as e:
        print(f"Erreur lors du test: {e}")
        import traceback

        traceback.print_exc()
