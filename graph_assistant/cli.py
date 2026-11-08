"""graph-assistant: the command line.

  graph-assistant login              sign in once (device code); delegated mode
  graph-assistant setup              create the folders, the workbook and the Excel table
  graph-assistant check              who am I, can I see the workbook and table
  graph-assistant run --dry-run      read new mail and show what would happen, change nothing
  graph-assistant run                do it
  graph-assistant watch --every 120  keep going
  graph-assistant report             what has been handled, spend, what needs a person
  graph-assistant logout             forget the cached sign-in

Settings: config/assistant.yaml (or --config). Identity comes from the environment:
  GRAPH_CLIENT_ID      the app registration's Application (client) ID (not a secret)
  GRAPH_TENANT         consumers (personal accounts, default), organizations, or your tenant id
  GRAPH_AUTH           device (default) or app (client credentials; needs GRAPH_CLIENT_SECRET)
  ANTHROPIC_API_KEY    for Claude
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time

import requests

from . import auth
from .assistant import Assistant, summarise
from .config import load
from .extract import ClaudeExtractor
from .graph import GraphClient, GraphError
from .store import Store


def _token(interactive: bool):
    client_id = os.environ.get("GRAPH_CLIENT_ID")
    if not client_id:
        sys.exit("GRAPH_CLIENT_ID is not set (the Application (client) ID from your Entra app registration)")
    tenant = os.environ.get("GRAPH_TENANT", "consumers")
    if os.environ.get("GRAPH_AUTH", "device") == "app":
        return auth.client_credentials_token(client_id, tenant)
    return auth.device_code_token(client_id, tenant, interactive=interactive)


def build(args, interactive=False) -> Assistant:
    cfg = load(args.config)
    graph = GraphClient(requests.Session(), _token(interactive), mailbox=cfg.mailbox, drive=cfg.drive)
    extractor = ClaudeExtractor(cfg.model) if args.cmd in ("run", "watch") else None
    return Assistant(graph, extractor, Store(args.state), cfg)


def main(argv=None):
    p = argparse.ArgumentParser(prog="graph-assistant")
    p.add_argument("--config", default="config/assistant.yaml")
    p.add_argument("--state", default=".graph-assistant/state.db")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("login")
    sub.add_parser("logout")
    sub.add_parser("setup")
    sub.add_parser("check")
    r = sub.add_parser("run")
    r.add_argument("--dry-run", action="store_true")
    w = sub.add_parser("watch")
    w.add_argument("--every", type=int, default=120, help="seconds between checks")
    sub.add_parser("report")
    args = p.parse_args(argv)

    try:
        if args.cmd == "logout":
            print("signed out" if auth.signed_out() else "was not signed in")
        elif args.cmd == "login":
            a = build(args, interactive=True)
            me = a.g.whoami()
            print(f"signed in as {me.get('mail') or me.get('userPrincipalName')}")
        elif args.cmd == "setup":
            print("workbook ready:", build(args).setup())
        elif args.cmd == "check":
            print(json.dumps(build(args).check(), indent=2))
        elif args.cmd == "run":
            print(summarise(build(args).run(dry_run=args.dry_run)))
        elif args.cmd == "watch":
            a = build(args)
            while True:
                outcomes = a.run()
                if outcomes:
                    print(time.strftime("%H:%M:%S"), summarise(outcomes), sep="\n", flush=True)
                time.sleep(args.every)
        elif args.cmd == "report":
            rows = Store(args.state).all()
            by = {}
            for r in rows:
                by[r["status"]] = by.get(r["status"], 0) + 1
            print(json.dumps(by), f"AI spend ${sum(r['cost_usd'] or 0 for r in rows):.4f}")
            for r in rows:
                if r["status"] in ("review", "error"):
                    steps = json.loads(r["steps"] or "{}")
                    print(f"- [{r['status']}] {r['sender']}: {r['subject']}  {r['last_error'] or ''} {steps.get('hold', '')}")
    except GraphError as e:
        sys.exit(f"Microsoft Graph said no: {e}")
    except auth.AuthError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
