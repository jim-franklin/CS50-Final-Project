from project import draft, greeting, writer_name, valediction


def main():
    test_draft()
    test_greeting()
    test_writer_name()
    valediction()


def test_draft():
    assert draft() == "DRAFT"


def test_greeting():
    assert greeting() == "Dear Sir/Madam,"


def test_valediction():
    assert valediction() == ("\n" * 10) + "Yours faithfully,\n"


def test_writer_name():
    assert writer_name() == "Ing. Oscar Amonoo-Neizer\n(Executive Secretary)"


if __name__ == "__main__":
    main()
