from __future__ import annotations

import argparse
import sys
from pathlib import Path

script_dir = Path(__file__).resolve().parent
repo_root = script_dir.parent
if str(script_dir) not in sys.path:
    sys.path.insert(0, str(script_dir))

from recommended_prices import run as recommended_prices_run, save_margin_settings
from update_prices import run as update_prices_run
from price_management import (
    action_add_to_actions,
    action_get_current_prices_and_actions,
    action_process_discount_requests,
    action_remove_unprofitable_actions,
    refresh_costs_views,
)
from utils import set_prompt_force


def cmd_save_margin(args: argparse.Namespace) -> int:
    if not (0 < args.min_margin < 1):
        raise ValueError("Минимальная маржа должна быть в диапазоне (0, 1).")
    if not (0 < args.desired_margin < 1):
        raise ValueError("Желаемая маржа должна быть в диапазоне (0, 1).")
    if args.min_margin >= args.desired_margin:
        raise ValueError("Минимальная маржа должна быть меньше желаемой.")
    save_margin_settings(repo_root, args.min_margin, args.desired_margin)
    print(
        "Настройки сохранены: "
        f"min_margin={args.min_margin:.2f}, desired_margin={args.desired_margin:.2f}"
    )
    return 0


def cmd_recommended_prices(args: argparse.Namespace) -> int:
    cmd_save_margin(args)
    recommended_prices_run(
        repo_root,
        min_margin=args.min_margin,
        desired_margin=args.desired_margin,
        include_actions=False,
    )
    return 0


def cmd_refresh_pricing(_args: argparse.Namespace) -> int:
    set_prompt_force(True)
    ok = action_get_current_prices_and_actions(repo_root)
    return 0 if ok else 1


def cmd_update_min_prices(_args: argparse.Namespace) -> int:
    update_prices_run(repo_root)
    return 0


def cmd_discount_requests(_args: argparse.Namespace) -> int:
    set_prompt_force(True)
    ok = action_process_discount_requests(repo_root)
    return 0 if ok else 1


def cmd_remove_unprofitable(_args: argparse.Namespace) -> int:
    set_prompt_force(True)
    ok = action_remove_unprofitable_actions(repo_root)
    if ok:
        refresh_costs_views(repo_root)
    return 0 if ok else 1


def cmd_add_to_actions(_args: argparse.Namespace) -> int:
    set_prompt_force(True)
    ok = action_add_to_actions(repo_root)
    if ok:
        refresh_costs_views(repo_root)
    return 0 if ok else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="UI helpers for OzonReportX.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    margin_parser = subparsers.add_parser("save-margin", help="Сохранить настройки маржи.")
    margin_parser.add_argument("--min-margin", type=float, required=True)
    margin_parser.add_argument("--desired-margin", type=float, required=True)
    margin_parser.set_defaults(handler=cmd_save_margin)

    recommended_parser = subparsers.add_parser(
        "recommended-prices",
        help="Сохранить маржу и рассчитать рекомендованные цены.",
    )
    recommended_parser.add_argument("--min-margin", type=float, required=True)
    recommended_parser.add_argument("--desired-margin", type=float, required=True)
    recommended_parser.set_defaults(handler=cmd_recommended_prices)

    refresh_parser = subparsers.add_parser(
        "refresh-pricing",
        help="Обновить текущие цены и активные акции в costs.xlsx.",
    )
    refresh_parser.set_defaults(handler=cmd_refresh_pricing)

    update_parser = subparsers.add_parser(
        "update-min-prices",
        help="Отправить минимальные цены в Ozon.",
    )
    update_parser.set_defaults(handler=cmd_update_min_prices)

    discount_parser = subparsers.add_parser(
        "discount-requests",
        help="Обработать заявки на скидку.",
    )
    discount_parser.set_defaults(handler=cmd_discount_requests)

    remove_parser = subparsers.add_parser(
        "remove-unprofitable-actions",
        help="Удалить невыгодные акции.",
    )
    remove_parser.set_defaults(handler=cmd_remove_unprofitable)

    add_parser = subparsers.add_parser(
        "add-to-actions",
        help="Добавить товары в акции.",
    )
    add_parser.set_defaults(handler=cmd_add_to_actions)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.handler(args))
    except Exception as exc:
        print(f"Ошибка: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
