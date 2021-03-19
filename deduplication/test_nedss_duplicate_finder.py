import pandas as pd
import pytest
import textdistance

import nedss_duplicate_finder as nedss


# this fixture is used for a couple basic test cases
# NOTE: the first and the last entry in this data is the same person
d = {'First Name': ["John", "Cindy", "Mario", "Albert", "Jon"],
     'Last Name ': ["McDermid", "McCain", "Battali", "Einstein", "McDermit"],
     'Address': ["35 Main St.", "42 Overlook Dr.", "103 Broadway", "88 Random House", "35 Mane Streat"],
     'Age': [42, 88, 23, 44, 41],
     'Gender': ["M", "F", "F", "M", "M"]}


def test_jaro_winkler():
    """Confirm that the jaro winkler implementation matches the original paper"""
    assert textdistance.jaro_winkler("campell", "campbell") == pytest.approx(0.9792, abs=0.01)
    assert textdistance.jaro_winkler("shakelford", "shakleford") == pytest.approx(0.9848, abs=0.01)
    assert textdistance.jaro_winkler("dwayne", "duane") == pytest.approx(0.84, abs=0.01)


def test_first_name_sim_nones():
    """Confirm that when a name is missing, it's considered a (trivial) match against other missing names"""
    assert nedss.first_name_similarity_scorer(None, None) == 1


def test_first_name_sim_empty_str():
    """Confirm that when a name is empty, it's considered a (trivial) match against other empty names"""
    assert nedss.first_name_similarity_scorer("", "") == 1


def test_address_sim():
    address = "123 Fake St., Faketown, FS, USA"
    assert nedss.address_similarity_scorer(address, address) == 1


def test_age_similarity_scorer_with_matching_ages():
    assert nedss.age_similarity_scorer("31", "31") == 0


def test_age_similarity_scorer_with_ages_within_two_year_tolerance():
    assert nedss.age_similarity_scorer("31", "33") == 0


def test_age_similarity_scorer_with_ages_outside_two_year_tolerance():
    assert nedss.age_similarity_scorer("31", "34") == -1


def test_age_similarity_is_trivial_match_when_age_is_not_an_int():
    assert nedss.age_similarity_scorer("", "34") == 0


def test_preprocess_phonetics():
    df = pd.DataFrame(data=d)
    df = nedss.preprocess(df)
    assert df.loc[0, "phonetic_first_name"] == "JN"
    assert df.loc[0, "phonetic_last_name"] == "MKTRMT"
    assert df.loc[0, "phonetic_address"] == "MN ST"


def test_score_aggregator_using_example_with_first_and_last_name_phonetic_match():
    df = pd.DataFrame(data=d)
    df = nedss.preprocess(df)
    df_first = df.loc[0]
    df_rest = df.loc[1:]
    df_scores = nedss.score_record_against_records(df_rest, df_first)
    assert df_scores.loc[4, "first_name_similarity_score"] == 1
    assert df_scores.loc[4, "last_name_similarity_score"] == 1
    assert df_scores.loc[4, "address_similarity_score"] == 1
    assert df_scores.loc[4, "gender_similarity_score"] == 0
    assert df_scores.loc[4, "age_similarity_score"] == 0
    assert df_scores.loc[4, "total_score"] == 3


def test_mini_pipeline():
    df = pd.DataFrame(data=d)
    df = nedss.preprocess(df)
    dupe_groups = nedss.get_dupe_index_groups(df, 3)
    assert dupe_groups == [[0, 4]]
