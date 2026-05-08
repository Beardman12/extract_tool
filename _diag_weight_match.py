import query_product_code as q
q._load_country_match_rules(None)
print('W1', q._weight_match('0<W≤0.2(50G起算)', '0-0.2（50g起步价）'))
print('W2', q._weight_match('0.2<W≤0.3', '0.201-0.3'))
print('W3', q._weight_match('0.3<W≤0.45', '0.301-0.45'))
print('W4', q._weight_match('0.45<W≤2', '0.451-2'))
print('C1', q._country_match('美国', 'US1,US2\n（财务务必两个都设置）'))
