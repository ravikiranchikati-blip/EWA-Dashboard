#!/usr/bin/env python3
"""
generate_dashboard.py v2
Reads ewa_data.json and injects it into the EWA dashboard template.
Template search order:
  1. ewa_template.html in the current working directory
  2. assets/dashboard.html in the skill directory
Usage: python generate_dashboard.py --input ewa_data.json --output ewa-dashboard.html
"""
import argparse, json, sys, os


def find_template():
    # 1. Check current working directory for ewa_template.html
    cwd_tpl = os.path.join(os.getcwd(), "ewa_template.html")
    if os.path.exists(cwd_tpl):
        print(f"[INFO] Using template: {cwd_tpl}", file=sys.stderr)
        with open(cwd_tpl, encoding="utf-8") as f:
            return f.read()
    # 2. Fall back to skill assets
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_dir  = os.path.dirname(script_dir)
    tpl_path   = os.path.join(skill_dir, "assets", "dashboard.html")
    if os.path.exists(tpl_path):
        print(f"[INFO] Using skill template: {tpl_path}", file=sys.stderr)
        with open(tpl_path, encoding="utf-8") as f:
            return f.read()
    print("ERROR: No template found. Expected ewa_template.html in working directory "
          f"or assets/dashboard.html in skill directory ({skill_dir}/assets/).", file=sys.stderr)
    sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Generate EWA HTML dashboard from JSON")
    parser.add_argument("--input",  "-i", required=True, help="Path to EWA JSON file")
    parser.add_argument("--output", "-o", required=True, help="Output HTML file path")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"ERROR: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    print(f"[INFO] Loading JSON: {args.input}", file=sys.stderr)
    with open(args.input, encoding="utf-8") as f:
        data = json.load(f)

    template = find_template()
    json_str = json.dumps(data, ensure_ascii=False)
    html = template.replace("__EWA_DATA__", json_str)

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Dashboard generated: {args.output}")


if __name__ == "__main__":
    main()
