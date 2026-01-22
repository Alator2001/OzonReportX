from __future__ import annotations

from typing import Optional

_PROMPT_FORCE: Optional[bool] = None


def set_prompt_force(value: Optional[bool]) -> None:
    global _PROMPT_FORCE
    _PROMPT_FORCE = value


def print_step(title: str):
    print(f"\n=== {title} ===")


def prompt_yes_no(message: str, default_yes: bool = True) -> bool:
    if _PROMPT_FORCE is not None:
        return _PROMPT_FORCE
    suffix = "[Y/n]" if default_yes else "[y/N]"
    while True:
        answer = input(f"{message} {suffix} ").strip().lower()
        if not answer:
            return default_yes
        if answer in ("y", "yes", "д", "да"):
            return True
        if answer in ("n", "no", "н", "нет"):
            return False
        print("Пожалуйста, ответьте 'y' или 'n'.")


