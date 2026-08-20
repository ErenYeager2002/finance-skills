# -*- coding: utf-8 -*-
"""fetch_zhiyun 纯函数单测（不连网、不碰账密）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

import fetch_zhiyun as F


def test_resolve_date_yesterday():
    d = F.resolve_date("yesterday")
    assert len(d) == 10 and d[4] == "-" and d[7] == "-"


def test_resolve_date_fixed():
    assert F.resolve_date("2026-07-21") == "2026-07-21"


def test_existing_exports_require_current_schema_version(tmp_path):
    day = "2026-07-27"
    stamp = "20260727"
    for role in ("回款记录", "订单交付", "核销明细", "订单明细"):
        (tmp_path / f"{role}_{stamp}.xlsx").write_bytes(b"fixture")

    # 旧四件套没有版本摘要时，默认禁止跳过重新取数。
    assert F.already_fetched(tmp_path, day) == []
    assert len(F.already_fetched(tmp_path, day, accept_unversioned=True)) == 4

    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        '{"export_schema_version":"old"}',
        encoding="utf-8",
    )
    assert F.already_fetched(tmp_path, day) == []

    (tmp_path / f"取数摘要_{stamp}.json").write_text(
        '{"export_schema_version":"' + F.EXPORT_SCHEMA_VERSION + '"}',
        encoding="utf-8",
    )
    assert len(F.already_fetched(tmp_path, day)) == 4


def test_plain_option_and_relation():
    opts = {"k1": "整笔回款"}
    assert F._plain('["k1"]', opts) == "整笔回款"
    assert F._plain('[{"name":"某某客户"}]') == "某某客户"
    assert F._plain(None) == ""


def test_settlement_relation_recovers_order_when_xiadan_is_empty():
    controls = [
        {"controlId": "order", "controlName": "结算订单"},
        {"controlId": "written", "controlName": "订单已核销金额"},
        {"controlId": "amount", "controlName": "交付额/原币"},
        {"controlId": "currency", "controlName": "结算币种"},
    ]
    rows = [{
        "order": '[{"name":"SO26000001"}]',
        "written": "120.47",
        "amount": "120.50",
        "currency": "人民币CNY",
    }]

    got = F.extract_related_orders(rows, controls, F.REL_JIESUAN)

    assert got == [{
        "so": "SO26000001",
        "written_off": "120.47",
        "written_off_local": "",
        "deliver": "120.50",
        "rate": "",
        "currency": "人民币CNY",
        "name": "",
        "delivery_date": "",
        "delivery_date_status": "",
        "source": "结算",
    }]


def test_related_order_reads_project_delivery_date_from_order_detail():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]
    rows = [{"order": "SO24100160", "date": "2025-08-13"}]

    got = F.extract_related_orders(rows, controls, F.REL_XIADAN)

    assert got[0]["delivery_date"] == "2025-08-13"
    assert got[0]["delivery_date_status"] == "关联下单明确值"


def test_lookup_order_delivery_date_requires_unique_explicit_date():
    controls = [
        {"controlId": "order", "controlName": "SO"},
        {"controlId": "date", "controlName": "项目交付日期"},
    ]

    class FakeClient:
        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        @staticmethod
        def search_rows(_worksheet_id, _so):
            return [
                {"order": "SO24100160", "date": "2025-08-13"},
                {"order": "SO24100160", "date": "2025-08-13"},
                {"order": "SO99999999", "date": "2024-01-01"},
            ]

    value, status = F.lookup_order_delivery_date(
        FakeClient(), "orders", controls, "SO24100160"
    )
    assert value == "2025-08-13"
    assert status == "订单详情明确值"


def test_no_credentials_in_source():
    src = Path(__file__).resolve().parents[1] / "scripts" / "fetch_zhiyun.py"
    text = src.read_text(encoding="utf-8")
    # 禁止真实账号/密码痕迹（允许文档里出现变量名 ZHIYUN_PASS）
    assert "sharon" not in text.lower()
    assert "sharon1234" not in text
    assert "getpass" not in text  # 核销任务不在取数中途交互询问
    assert "input(" not in text
    assert "自动任务不会在取数过程中弹出账号密码询问" in text
    # 禁止把真实密码字面量赋给环境示例
    assert "PASS='****'" not in text


def test_historical_writeoffs_for_sos_gets_cross_parent_history_only():
    names = [
        "核销记录NUM", "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
        "核销日期", "币种", "汇率", "订单名称", "是否已撤销",
    ]
    controls = [
        {"controlId": f"c{i}", "controlName": name}
        for i, name in enumerate(names)
    ]

    def row(**values):
        return {
            f"c{i}": values.get(name, "")
            for i, name in enumerate(names)
        }

    class FakeClient:
        def controls(self, _worksheet_id):
            return controls

        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        def search_rows(self, _worksheet_id, so):
            assert so == "SO26000001"
            return [
                row(
                    核销记录NUM="HX_OLD_001",
                    回款记录NUM='[{"name":"AR_OLD_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=30,
                    本次核销金额本币=30,
                    核销日期="2026-06-11",
                ),
                row(  # 目标日当前行由父回款关联子表负责，不在这里重复补。
                    核销记录NUM="HX_NOW_001",
                    回款记录NUM='[{"name":"AR_NOW_001"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=10,
                    本次核销金额本币=10,
                    核销日期="2026-07-24",
                ),
                row(  # 已撤销历史行不计。
                    核销记录NUM="HX_OLD_002",
                    回款记录NUM='[{"name":"AR_OLD_002"}]',
                    订单NUM='[{"name":"SO26000001"}]',
                    本次核销金额=5,
                    本次核销金额本币=5,
                    核销日期="2026-05-01",
                    是否已撤销="是",
                ),
                row(  # 全文搜索误命中的其它 SO 必须精确排除。
                    核销记录NUM="HX_OTHER",
                    回款记录NUM='[{"name":"AR_OTHER"}]',
                    订单NUM='[{"name":"SO260000010"}]',
                    本次核销金额=99,
                    本次核销金额本币=99,
                    核销日期="2026-04-01",
                ),
            ]

    got = F.historical_writeoffs_for_sos(
        FakeClient(), "WS_MX", ["SO26000001"], "2026-07-24"
    )
    assert len(got) == 2  # 撤销记录也保留给分类器审计，但不参与金额。
    assert got[0][0] == "HX_OLD_001"
    assert got[0][2] == "AR_OLD_001"
    assert got[0][3] == "2026-06-11"
    assert got[0][8] == "SO26000001"
    assert got[1][10] == "是"


def test_historical_writeoffs_only_dedup_same_record_id_and_never_business_fields():
    names = [
        "核销记录NUM", "回款记录NUM", "订单NUM", "本次核销金额", "本次核销金额本币",
        "核销日期", "币种", "汇率", "订单名称", "是否已撤销",
    ]
    controls = [
        {"controlId": f"c{i}", "controlName": name}
        for i, name in enumerate(names)
    ]

    def row(record_id):
        values = {
            "核销记录NUM": record_id,
            "回款记录NUM": '[{"name":"AR26070001"}]',
            "订单NUM": '[{"name":"SO26000001"}]',
            "本次核销金额": 100,
            "本次核销金额本币": 100,
            "核销日期": "2026-07-01",
            "币种": "CNY",
        }
        return {
            **{f"c{i}": values.get(name, "") for i, name in enumerate(names)},
            "rowid": f"ROW-{record_id or 'NONE'}",
        }

    class FakeClient:
        def controls(self, _worksheet_id):
            return controls

        @staticmethod
        def name_map(ctrls):
            return {c["controlId"]: c["controlName"] for c in ctrls}

        @staticmethod
        def option_maps(_ctrls):
            return {}

        def search_rows(self, _worksheet_id, _so):
            return [row("HX1"), row("HX1"), row("HX2"), row(""), row("")]

    got = F.historical_writeoffs_for_sos(
        FakeClient(), "WS_MX", ["SO26000001"], "2026-07-31"
    )
    assert [item[0] for item in got] == ["HX1", "HX2", "", ""]


def test_supplement_identifiers_are_searched_exactly_and_reported_against_exports():
    class FakeClient:
        @staticmethod
        def controls(_worksheet_id):
            return [
                {"controlId": "order", "controlName": "SO"},
                {"controlId": "ar", "controlName": "回款记录NUM"},
            ]

        @staticmethod
        def id_by_name(_controls, _name):
            return "relation"

        @staticmethod
        def datasource_of(_worksheet_id, relation):
            return {
                F.REL_XIADAN: "orders",
                F.REL_HEXIAO_MINGXI: "writeoffs",
                F.REL_SODLINE: "details",
            }.get(relation, "")

        @staticmethod
        def search_rows(worksheet_id, identifier):
            if worksheet_id == F.WS_HUIKUAN and identifier == "AR26070140":
                return [
                    {F.F_HK["ar"]: "AR26070140"},
                    {F.F_HK["ar"]: "AR260701400"},
                ]
            if worksheet_id == "orders" and identifier == "SO26020320":
                return [{"order": '[{"name":"SO26020320"}]'}]
            return []

    searched = F.search_supplement_identifiers(
        FakeClient(),
        ["AR26070140", "AR26079999"],
        ["SO26020320", "SO26029999"],
    )
    assert searched == {
        "found_ar_ids": ["AR26070140"],
        "found_so_ids": ["SO26020320"],
    }

    result = F.build_supplement_result(
        ["AR26070140", "AR26079999"],
        ["SO26020320", "SO26029999"],
        before={"ar_ids": set(), "so_ids": {"SO26020320"}},
        after={"ar_ids": {"AR26070140"}, "so_ids": {"SO26020320"}},
        searched=searched,
    )
    assert result["added"]["ar_ids"] == ["AR26070140"]
    assert result["existing"]["so_ids"] == ["SO26020320"]
    assert result["unresolved"]["ar_ids"] == ["AR26079999"]
    assert result["unresolved"]["so_ids"] == ["SO26029999"]


def test_resolve_fetch_dates_keeps_workdays_in_chronological_order():
    assert F.resolve_fetch_dates(
        single_date="",
        date_from="2026-08-14",
        date_to="2026-08-18",
        include_weekends=False,
    ) == ["2026-08-14", "2026-08-17", "2026-08-18"]


def test_date_range_fetch_logs_in_once_and_writes_each_day_separately(
    tmp_path, monkeypatch
):
    login_calls = []
    client_calls = []
    filtered_days = []

    def fake_login(base_url, user, password, *, headless):
        login_calls.append((base_url, user, password, headless))
        return "session-cookie", "account-id"

    class FakeClient:
        def __init__(self, base_url, cookie, *, account_id=""):
            client_calls.append((base_url, cookie, account_id))

        @staticmethod
        def controls(_worksheet_id):
            return [
                {
                    "controlId": "xiadan",
                    "controlName": F.REL_XIADAN,
                    "dataSource": "orders",
                }
            ]

        @staticmethod
        def option_maps(_controls):
            return {}

        @staticmethod
        def id_by_name(controls, name):
            for control in controls:
                if control.get("controlName") == name:
                    return control.get("controlId", "")
            return ""

        @staticmethod
        def datasource_of(_worksheet_id, relation):
            return "orders" if relation == F.REL_XIADAN else ""

        @staticmethod
        def filter_rows_by_date(_worksheet_id, _control_id, day):
            filtered_days.append(day)
            return [], 0

    monkeypatch.delenv("MD_PSS_ID", raising=False)
    monkeypatch.setattr(F, "resolve_credentials", lambda _args: ("user", "secret"))
    monkeypatch.setattr(F, "login_with_password", fake_login)
    monkeypatch.setattr(F, "ZhiyunClient", FakeClient)

    assert F.main([
        "--date-from", "2026-08-14",
        "--date-to", "2026-08-17",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 0

    assert len(login_calls) == 1
    assert len(client_calls) == 1
    assert filtered_days == ["2026-08-14", "2026-08-17"]
    export_dir = tmp_path / "01_智云导出"
    for tag in ("20260814", "20260817"):
        assert (export_dir / f"回款记录_{tag}.xlsx").is_file()
        assert (export_dir / f"订单交付_{tag}.xlsx").is_file()
        assert (export_dir / f"核销明细_{tag}.xlsx").is_file()
        assert (export_dir / f"订单明细_{tag}.xlsx").is_file()
        assert (export_dir / f"取数摘要_{tag}.json").is_file()


def test_date_range_uses_current_daily_exports_without_logging_in(
    tmp_path, monkeypatch
):
    export_dir = tmp_path / "01_智云导出"
    export_dir.mkdir()
    for day in ("2026-08-14", "2026-08-17"):
        stamp = day.replace("-", "")
        for role in ("回款记录", "订单交付", "核销明细", "订单明细"):
            (export_dir / f"{role}_{stamp}.xlsx").write_bytes(b"fixture")
        (export_dir / f"取数摘要_{stamp}.json").write_text(
            '{"export_schema_version":"' + F.EXPORT_SCHEMA_VERSION + '"}',
            encoding="utf-8",
        )

    def unexpected_login(*_args, **_kwargs):
        raise AssertionError("全部日期命中缓存时不应登录智云")

    monkeypatch.delenv("MD_PSS_ID", raising=False)
    monkeypatch.setattr(F, "resolve_credentials", unexpected_login)

    assert F.main([
        "--date-from", "2026-08-14",
        "--date-to", "2026-08-17",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 0


def test_date_range_stops_after_a_daily_fetch_failure(tmp_path, monkeypatch):
    filtered_days = []

    class FakeClient:
        def __init__(self, *_args, **_kwargs):
            pass

        @staticmethod
        def controls(_worksheet_id):
            return [{
                "controlId": "xiadan",
                "controlName": F.REL_XIADAN,
                "dataSource": "orders",
            }]

        @staticmethod
        def option_maps(_controls):
            return {}

        @staticmethod
        def id_by_name(controls, name):
            for control in controls:
                if control.get("controlName") == name:
                    return control.get("controlId", "")
            return ""

        @staticmethod
        def datasource_of(_worksheet_id, relation):
            return "orders" if relation == F.REL_XIADAN else ""

        @staticmethod
        def filter_rows_by_date(_worksheet_id, _control_id, day):
            filtered_days.append(day)
            if day == "2026-08-17":
                raise RuntimeError("simulated daily fetch failure")
            return [], 0

    monkeypatch.setenv("MD_PSS_ID", "session-cookie")
    monkeypatch.setattr(F, "ZhiyunClient", FakeClient)

    assert F.main([
        "--date-from", "2026-08-14",
        "--date-to", "2026-08-18",
        "--workspace", str(tmp_path),
        "--skip-gap-check",
    ]) == 2
    assert filtered_days == ["2026-08-14", "2026-08-17"]
    assert not (
        tmp_path / "01_智云导出" / "取数摘要_20260818.json"
    ).exists()


def test_date_range_rejects_identifier_supplement(tmp_path):
    import pytest

    with pytest.raises(SystemExit):
        F.main([
            "--date-from", "2026-08-14",
            "--date-to", "2026-08-17",
            "--supplement-ar", "AR26080001",
            "--workspace", str(tmp_path),
            "--skip-gap-check",
        ])
