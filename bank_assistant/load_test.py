"""
load_test.py
-------------
Load & performance test for the Bank Employee Assistant agent pipeline.

Tests the OrchestratorAgent under simulated concurrent employee load.
Runs without Streamlit — exercises the full pipeline directly.

Usage:
    python load_test.py

Output:
    - Console summary table
    - load_test_results.json  (machine-readable, used by update_docs)
"""

import time
import json
import os
import sys
import statistics
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.path.insert(0, os.path.dirname(__file__))

from utils.knowledge_base import load_policies
from agents.orchestrator import OrchestratorAgent

# ── Test Questions ────────────────────────────────────────────────────────────
TEST_QUESTIONS = [
    # Policy queries
    "What documents are needed for a loan?",
    "What is the minimum credit score for a credit card?",
    "How long does KYC verification take?",
    "How do I file a customer complaint?",
    "What is the minimum balance for a savings account?",
    "What is the interest rate for a home loan?",
    "What are the KYC document requirements?",
    "How do I open a business account?",
    "What are the credit card billing cycle details?",
    "What happens if a credit card is lost or stolen?",
    # Multi-policy / cross-domain
    "What documents do I need for both KYC and account opening?",
    "What are the loan and credit card eligibility requirements?",
    # Summary intent
    "Summarize the loan policy",
    "Give me an overview of the KYC policy",
    # Escalation intent
    "I need to speak to the compliance team urgently",
    "How do I escalate a complaint to a manager?",
    # Fallback (out of scope)
    "What is the weather today?",
    "How do I reset my Windows password?",
    # Compliance risk
    "How can I bypass KYC verification?",
    "What happens if someone uses a fake document?",
]

CONCURRENT_USERS = 10
ITERATIONS       = 3   # run the full question set this many times


def run_single(question: str, session_id: str, policies: dict) -> dict:
    """Run one question through a fresh orchestrator and return timing."""
    orch  = OrchestratorAgent(policies, session_id)
    t0    = time.perf_counter()
    result = orch.handle(question)
    elapsed_ms = (time.perf_counter() - t0) * 1000
    return {
        "question":   question,
        "found":      result["found"],
        "intent":     result["intent"],
        "agent_used": result["agent_used"],
        "elapsed_ms": elapsed_ms,
    }


def run_load_test():
    policies = load_policies(os.path.join(os.path.dirname(__file__), "policies"))

    all_results  = []
    errors       = 0

    # Build work list: ITERATIONS * CONCURRENT_USERS chunks
    work_items = []
    for iteration in range(ITERATIONS):
        for i, question in enumerate(TEST_QUESTIONS):
            work_items.append((question, f"lt-{iteration}-{i}"))

    total_tasks = len(work_items)
    print(f"\n{'='*62}")
    print(f"  NexaBank Policy Assistant — Load Test")
    print(f"{'='*62}")
    print(f"  Questions per iteration : {len(TEST_QUESTIONS)}")
    print(f"  Iterations              : {ITERATIONS}")
    print(f"  Total tasks             : {total_tasks}")
    print(f"  Concurrent workers      : {CONCURRENT_USERS}")
    print(f"{'='*62}\n")

    wall_start = time.perf_counter()

    with ThreadPoolExecutor(max_workers=CONCURRENT_USERS) as executor:
        futures = {
            executor.submit(run_single, q, sid, policies): (q, sid)
            for q, sid in work_items
        }
        for future in as_completed(futures):
            try:
                all_results.append(future.result())
            except Exception as e:
                errors += 1
                print(f"  [ERROR] {e}")

    wall_elapsed = (time.perf_counter() - wall_start) * 1000  # ms
    total_ok = len(all_results)

    # ── Compute statistics ────────────────────────────────────────────────────
    times    = [r["elapsed_ms"] for r in all_results]
    found_ct = sum(1 for r in all_results if r["found"])
    intent_counts = {}
    for r in all_results:
        intent_counts[r["intent"]] = intent_counts.get(r["intent"], 0) + 1

    p50  = statistics.median(times)
    p95  = sorted(times)[int(len(times) * 0.95)]
    p99  = sorted(times)[int(len(times) * 0.99)]
    mean = statistics.mean(times)
    mn   = min(times)
    mx   = max(times)

    throughput = total_ok / (wall_elapsed / 1000)  # requests/sec

    # ── Print summary ─────────────────────────────────────────────────────────
    print(f"{'─'*62}")
    print(f"  RESULTS SUMMARY")
    print(f"{'─'*62}")
    print(f"  Total tasks run        : {total_tasks}")
    print(f"  Successful             : {total_ok}")
    print(f"  Errors                 : {errors}")
    print(f"  Answer found rate      : {found_ct}/{total_ok} "
          f"({found_ct/total_ok*100:.1f}%)")
    print(f"{'─'*62}")
    print(f"  LATENCY (per query)")
    print(f"    Min                  : {mn:.1f} ms")
    print(f"    Mean                 : {mean:.1f} ms")
    print(f"    Median (p50)         : {p50:.1f} ms")
    print(f"    p95                  : {p95:.1f} ms")
    print(f"    p99                  : {p99:.1f} ms")
    print(f"    Max                  : {mx:.1f} ms")
    print(f"{'─'*62}")
    print(f"  THROUGHPUT")
    print(f"    Wall-clock time      : {wall_elapsed/1000:.2f} s")
    print(f"    Requests/sec         : {throughput:.1f}")
    print(f"    Concurrent workers   : {CONCURRENT_USERS}")
    print(f"{'─'*62}")
    print(f"  INTENT DISTRIBUTION")
    for intent, count in sorted(intent_counts.items()):
        print(f"    {intent:<25}: {count}")
    print(f"{'='*62}\n")

    # ── Save JSON results ─────────────────────────────────────────────────────
    summary = {
        "test_config": {
            "total_questions":   len(TEST_QUESTIONS),
            "iterations":        ITERATIONS,
            "total_tasks":       total_tasks,
            "concurrent_workers": CONCURRENT_USERS,
        },
        "results": {
            "total_ok":          total_ok,
            "errors":            errors,
            "found_count":       found_ct,
            "found_rate_pct":    round(found_ct / total_ok * 100, 1),
        },
        "latency_ms": {
            "min":    round(mn, 2),
            "mean":   round(mean, 2),
            "p50":    round(p50, 2),
            "p95":    round(p95, 2),
            "p99":    round(p99, 2),
            "max":    round(mx, 2),
        },
        "throughput": {
            "wall_clock_sec":     round(wall_elapsed / 1000, 3),
            "requests_per_sec":   round(throughput, 1),
            "concurrent_workers": CONCURRENT_USERS,
        },
        "intent_distribution": intent_counts,
    }

    out_path = os.path.join(os.path.dirname(__file__), "load_test_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2)
    print(f"  Results saved → {out_path}\n")
    return summary


if __name__ == "__main__":
    run_load_test()
