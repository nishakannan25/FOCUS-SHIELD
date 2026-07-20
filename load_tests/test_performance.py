"""
FOCUS-SHIELD – Performance Load Testing Simulator
==================================================
Simulates concurrent web traffic targeting the Vercel-hosted production backend
and measures latency / success rates.
"""

import time
import urllib.request
import concurrent.futures

TARGET_URL = "https://focus-shield-three.vercel.app"
CONCURRENT_WORKERS = 5
REQUESTS_PER_WORKER = 2

def ping_target():
    start = time.time()
    try:
        with urllib.request.urlopen(TARGET_URL, timeout=10) as response:
            code = response.getcode()
            latency = (time.time() - start) * 1000
            return code, latency
    except Exception as e:
        return 500, 0.0

def test_load_performance():
    print(f"[INFO] Initiating concurrent connection test to {TARGET_URL}")
    results = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=CONCURRENT_WORKERS) as executor:
        futures = [executor.submit(ping_target) for _ in range(CONCURRENT_WORKERS * REQUESTS_PER_WORKER)]
        for fut in concurrent.futures.as_completed(futures):
            code, latency = fut.result()
            results.append((code, latency))

    successes = [r for r in results if r[0] == 200]
    latencies = [r[1] for r in results if r[0] == 200]

    success_rate = (len(successes) / len(results)) * 100
    avg_latency = sum(latencies) / len(latencies) if latencies else 0.0

    print(f"[INFO] Load Simulation Complete. Success Rate: {success_rate:.2f}%, Avg Latency: {avg_latency:.2f}ms")
    
    # Validation assertions
    assert success_rate >= 80.0, f"Error rate too high: {100 - success_rate}%"
    assert avg_latency < 5000.0, f"Average response latency exceeded maximum threshold: {avg_latency}ms"

if __name__ == "__main__":
    test_load_performance()
