import pytest

from sokolov.openai_assets import is_obvious_case

def test_is_obvious_case():
    # Test 1: Expression "all of them"
    body = "We invited ten friends, and all of them came."
    span = (35, 39)  # The span of 'them'
    assert is_obvious_case(body, span) == True

    # Test 2: No special expression
    body = "They don't like this."
    span = (0, 4)
    assert is_obvious_case(body, span) == False

    # Test 3: Expression "they were all"
    body = "They were all excited about the trip."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 4: Expression "they are all"
    body = "They are all coming to the party."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 5: Expression "they will all"
    body = "They will all be there at the meeting."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 6: Expression "some of them"
    body = "Some of them were late to the event."
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    # Test 7: No special expression, similar structure
    body = "Most will arrive by noon."
    span = (0, 4)
    assert is_obvious_case(body, span) == False

    body = "There are more than just three of them, you know"
    span= (34, 38)
    assert is_obvious_case(body, span) == True

    body = "They resemble each other"
    span = (0, 4)
    assert is_obvious_case(body, span) == True

    body = "Last time I check they loved one another."
    span = (18, 22)
    assert is_obvious_case(body, span) == True

    body = '''I think  they have a new site up, they all should just go there and be done with it.'''
    span = (34, 38)
    assert is_obvious_case(body, span) == True

    body = "All of them!"
    span = (7, 11)
    assert is_obvious_case(body, span) == True