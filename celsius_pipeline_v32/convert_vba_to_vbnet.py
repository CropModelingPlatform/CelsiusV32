#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path


DEFAULT_STANDARD_MODULES = {"Functions"}
DEFAULT_SKIP_MODULES = {
    "Principal",
    "NumeroteEnreg",
    "GraphesJournaliers",
    "Compteur_param",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert exported Access/VBA modules into generated VB .NET files.")
    parser.add_argument("export_dir", type=Path, help="Directory containing exported .bas files.")
    parser.add_argument("project_dir", type=Path, help="Target VB .NET project directory.")
    parser.add_argument(
        "--converted-subdir",
        default="Converted",
        help="Subdirectory inside the project receiving generated .vb files.",
    )
    parser.add_argument(
        "--standard-module",
        action="append",
        default=[],
        help="Module names that should stay as VB Modules instead of Classes.",
    )
    parser.add_argument(
        "--skip-module",
        action="append",
        default=[],
        help="Module names to ignore during generation.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    standard_modules = DEFAULT_STANDARD_MODULES.union(args.standard_module)
    skip_modules = DEFAULT_SKIP_MODULES.union(args.skip_module)
    dst = args.project_dir / args.converted_subdir
    dst.mkdir(parents=True, exist_ok=True)

    for existing in dst.glob("*.vb"):
        existing.unlink()

    for src_file in sorted(args.export_dir.glob("*.bas")):
        name = src_file.stem
        if name in skip_modules:
            continue

        text = src_file.read_text(encoding="latin-1")
        converted = convert_file(name, text, standard_modules)
        (dst / f"{name}.vb").write_text(converted, encoding="utf-8")


def convert_file(name: str, text: str, standard_modules: set[str]) -> str:
    text = text.replace("\r\n", "\n")
    text = re.sub(r"^Attribute .*$\n?", "", text, flags=re.MULTILINE)
    text = text.replace("Option Compare Database", "")
    text = text.replace("Option Explicit", "")

    text = convert_special_properties(text)
    text = convert_property_gets(text)
    text = re.sub(r"^\s*MsgBox\s*(.+)$", r"Console.Error.WriteLine(\1)", text, flags=re.MULTILINE)
    text = re.sub(r"(?<!End )\bSet\s+", "", text)
    text = re.sub(r"(\b[A-Za-z_][A-Za-z0-9_]*)!([A-Za-z_][A-Za-z0-9_]*)", r'\1("\2")', text)
    text = text.replace("Variant", "Object")
    text = re.sub(r"\bNull\b", "Nothing", text)
    text = re.sub(r"As\s+String\s+\*\s+\d+", "As String", text)
    text = re.sub(r"(\.\s*Open)\s+(.+)$", r"\1(\2)", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*Wend\s*$", "End While", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*DoEvents\s*$", "System.Threading.Thread.Yield()", text, flags=re.MULTILINE)
    text = re.sub(r"^\s*(Sub|Function)\s+", r"Public \1 ", text, flags=re.MULTILINE)
    text = convert_array_bounds(text)

    header = [
        "Option Strict Off",
        "Option Explicit Off",
        "Imports System",
        "",
    ]

    body = text.strip() + "\n"
    body = "\n".join(convert_declaration_line(line) for line in body.splitlines()) + "\n"

    if name in standard_modules:
        wrapped = f"Public Module {name}\n{indent(body)}End Module\n"
    else:
        wrapped = f"Public Class {name}\n{indent(body)}End Class\n"

    return "\n".join(header) + wrapped


def convert_special_properties(text: str) -> str:
    replacements = {
        r"Property Get nNbJourSimul\(\) As Integer\s+nNbJourSimul = NbJourSimul\s+End Property": """
Public Property nNbJourSimul As Integer
    Get
        Return NbJourSimul
    End Get
    Set(value As Integer)
        NbJourSimul = value
    End Set
End Property
""",
        r"Property Let nNbJourSimul\(NJsimul As Integer\)\s+NbJourSimul = NJsimul\s+End Property": "",
        r"Property Get ncropsta\(icult As Integer, joursim As Integer\) As Integer\s+ncropsta = cropsta\(icult, joursim\)\s+End Property": """
Public Property ncropsta(icult As Integer, joursim As Integer) As Integer
    Get
        Return cropsta(icult, joursim)
    End Get
    Set(value As Integer)
        cropsta(icult, joursim) = value
    End Set
End Property
""",
        r"Property Let ncropsta\(icult As Integer, joursim As Integer, CSTA As Integer\)\s+cropsta\(icult, joursim\) = CSTA\s+End Property": "",
        r"Property Get bDie\(icult As Integer\) As Boolean\s+bDie = Die\(icult\)\s+End Property": """
Public Property bDie(icult As Integer) As Boolean
    Get
        Return Die(icult)
    End Get
    Set(value As Boolean)
        Die(icult) = value
    End Set
End Property
""",
        r"Property Let bDie\(icult As Integer, FinCult As Boolean\)\s+Die\(icult\) = FinCult\s+End Property": "",
    }

    for pattern, replacement in replacements.items():
        text = re.sub(
            pattern,
            replacement.strip() + "\n",
            text,
            flags=re.MULTILINE | re.DOTALL | re.IGNORECASE,
        )

    return text


def convert_property_gets(text: str) -> str:
    pattern = re.compile(
        r"Property Get (\w+)\((.*?)\)(?: As ([^\n]+))?\n(.*?)End Property",
        flags=re.DOTALL,
    )

    def repl(match: re.Match[str]) -> str:
        name = match.group(1)
        args = match.group(2)
        return_type = (match.group(3) or "Object").strip()
        body = match.group(4).rstrip()
        body = re.sub(rf"^\s*{re.escape(name)}\s*=\s*", "Return ", body, flags=re.MULTILINE)
        return f"Public Function {name}({args}) As {return_type}\n{body}\nEnd Function"

    return pattern.sub(repl, text)


def convert_array_bounds(text: str) -> str:
    pattern = re.compile(r"^(\s*(?:Dim|Public|Private)\s+\w+\()([^)]+)(\).*)$", flags=re.MULTILINE)

    def repl(match: re.Match[str]) -> str:
        prefix, bounds, suffix = match.groups()
        dims = [re.sub(r"^\s*-?\d+\s+To\s+", "", dim.strip(), flags=re.IGNORECASE) for dim in bounds.split(",")]
        return prefix + ", ".join(dims) + suffix

    return pattern.sub(repl, text)


def convert_declaration_line(line: str) -> str:
    line = line.rstrip()
    match = re.match(r"^(\s*(?:Public\s+)?(?:Sub|Function)\s+\w+\()([^)]+)(\).*)$", line)
    if not match:
        return line

    prefix, params, suffix = match.groups()
    items = [item.strip() for item in params.split(",")]
    if any(" As " in item for item in items):
        fixed = [item if " As " in item else f"{item} As Object" for item in items]
        return prefix + ", ".join(fixed) + suffix

    return line


def indent(text: str) -> str:
    return "".join(("    " + line if line.strip() else line) + "\n" for line in text.splitlines())


if __name__ == "__main__":
    main()
