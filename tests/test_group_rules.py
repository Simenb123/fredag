from fredag.group_rules import GroupRule, resolve_group

def test_resolve_group_domain():
    rules = [GroupRule(name="KundeX", target_dir=".", senders=["@kundex.no"])]
    r = resolve_group(rules, "sb@kundex.no", "S B")
    assert r and r.name == "KundeX"

def test_resolve_group_email_exact():
    rules = [GroupRule(name="VIP", target_dir=".", senders=["vip@firma.no"])]
    r = resolve_group(rules, "vip@firma.no", "VIP")
    assert r and r.name == "VIP"

def test_resolve_group_wildcard_in_name():
    rules = [GroupRule(name="Leverandør", target_dir=".", senders=["*as"]) ]
    r = resolve_group(rules, "no-reply@annet.no", "Fabrikk AS")
    assert r and r.name == "Leverandør"
