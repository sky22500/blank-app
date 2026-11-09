calc_krw, calc_krw_vn, calc_krw_os = 0.0, 0.0, 0.0
        for item in carryover_list:
            q = item['qty']
            gh_clean = clean_key(item['gh'])
            gh_suf = get_gh_suffix(item['gh'])
            b_mod = item['model'][:-1] if item['model'].endswith('E') else item['model']
            mod_clean = clean_key(b_mod)
            
            p = gh_map.get(gh_clean, 0.0)
            if p == 0: p = suffix_map.get(gh_suf, 0.0)
            if p == 0: p = model_map.get(mod_clean, 0.0)
            if p == 0:
                for mk, mv in gh_map.items():
                    if len(mk) >= 5 and (mk in gh_clean or gh_clean in mk or gh_suf in mk):
                        p = mv; break
            if p == 0: p = avg_price
            
            i_krw = q * p
            calc_krw += i_krw
            if is_vn_site(item['site']): calc_krw_vn += i_krw
            else: calc_krw_os += i_krw
            
            # 💡 [추가된 로직] 이월 수량과 금액을 모델별, SITE별 실적(model_data, site_data)에 직접 더해줍니다.
            rev_ok = i_krw / 100_000_000.0
            if b_mod:
                if b_mod not in model_data: model_data[b_mod] = {'Actual': 0, 'Rev_Act': 0.0}
                model_data[b_mod]['Actual'] += q
                model_data[b_mod]['Rev_Act'] += rev_ok
            if item['site']:
                if item['site'] not in site_data: site_data[item['site']] = {'Actual': 0, 'Rev_Act': 0.0}
                site_data[item['site']]['Actual'] += q
                site_data[item['site']]['Rev_Act'] += rev_ok
            
        total_carryover_krw = calc_krw
