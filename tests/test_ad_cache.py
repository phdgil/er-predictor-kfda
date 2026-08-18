import pickle
import tempfile
import threading
import unittest
from pathlib import Path

import numpy as np

from core.ad import ADCalculator


class ADCacheTests(unittest.TestCase):
    def _fit_calculator(self, reference_path):
        calculator = ADCalculator(n_neighbors=2)
        calculator.fit(
            np.array([[0.0, 1.0, 0.0], [1.0, 0.0, 1.0], [1.0, 1.0, 0.0]]),
            reference_path=str(reference_path),
            metadata=[{"Name": "one"}, {"Name": "two"}, {"Name": "three"}],
        )
        return calculator

    def test_legacy_pickle_is_rejected_without_execution(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            reference = directory / "reference.xlsx"
            reference.write_bytes(b"reference")
            marker = directory / "executed"

            class Malicious:
                def __reduce__(self):
                    return (Path.touch, (marker,))

            cache = directory / "legacy.pkl"
            cache.write_bytes(pickle.dumps(Malicious()))
            with self.assertRaisesRegex(ValueError, "legacy pickle"):
                ADCalculator(n_neighbors=2).load_cache(str(cache), str(reference))
            self.assertFalse(marker.exists())

    def test_malformed_cache_and_schema_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            reference = directory / "reference.xlsx"
            reference.write_bytes(b"reference")
            malformed = directory / "malformed.npz"
            malformed.write_bytes(b"not an npz cache")
            with self.assertRaises(ValueError):
                ADCalculator(n_neighbors=2).load_cache(str(malformed), str(reference))

            incomplete = directory / "incomplete.npz"
            np.savez_compressed(incomplete, metadata=np.asarray("{}"))
            with self.assertRaisesRegex(ValueError, "schema"):
                ADCalculator(n_neighbors=2).load_cache(str(incomplete), str(reference))

    def test_cache_is_bound_to_reference_content_not_size(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            reference = directory / "reference.xlsx"
            reference.write_bytes(b"aaaa")
            cache = directory / "ad-cache.npz"
            self._fit_calculator(reference).save_cache(str(cache))
            reference.write_bytes(b"bbbb")
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                ADCalculator(n_neighbors=2).load_cache(str(cache), str(reference))

    def test_safe_cache_round_trip_restores_predictions(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            reference = directory / "reference.xlsx"
            reference.write_bytes(b"reference")
            cache = directory / "ad-cache.npz"
            original = self._fit_calculator(reference)
            expected = original.predict(np.array([[0.0, 1.0, 0.0]]))
            original.save_cache(str(cache))

            restored = ADCalculator(n_neighbors=2).load_cache(str(cache), str(reference))
            actual = restored.predict(np.array([[0.0, 1.0, 0.0]]))
            self.assertEqual(actual.in_domain, expected.in_domain)
            self.assertAlmostEqual(actual.distance, expected.distance)
            self.assertAlmostEqual(actual.threshold, expected.threshold)

    def test_adjacent_safe_cache_bootstraps_writable_state_cache(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            reference = directory / "reference.xlsx"
            reference.write_bytes(b"reference")
            adjacent_cache = directory / "reference_ad_cache.npz"
            state_cache_root = directory / "state-cache"
            expected = self._fit_calculator(reference)
            expected.save_cache(str(adjacent_cache))

            restored = ADCalculator(n_neighbors=2, cache_root=str(state_cache_root))

            def refit_should_not_run(_path):
                raise AssertionError("adjacent safe cache should load before refitting")

            restored.fit_from_excel = refit_should_not_run
            restored.fit_from_excel_cached(str(reference))

            actual = restored.predict(np.array([[0.0, 1.0, 0.0]]))
            self.assertTrue(restored.fitted)
            self.assertAlmostEqual(actual.distance, expected.predict(np.array([[0.0, 1.0, 0.0]])).distance)
            self.assertTrue(Path(restored.default_cache_path(str(reference))).is_file())
            self.assertIn(f"AD cache loaded: {adjacent_cache}", restored.last_cache_status)

    def test_cache_invalidation_diagnostic_survives_refit_and_save(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            reference = directory / "reference.xlsx"
            reference.write_bytes(b"reference")
            cache = directory / "ad-cache.npz"
            cache.write_bytes(b"invalid")
            calculator = ADCalculator(n_neighbors=2)

            def refit(path):
                return calculator.fit(
                    np.array([[0.0, 1.0, 0.0], [1.0, 0.0, 1.0], [1.0, 1.0, 0.0]]),
                    reference_path=path,
                )

            calculator.fit_from_excel = refit
            calculator.fit_from_excel_cached(str(reference), str(cache))
            self.assertIn("AD cache not used:", calculator.last_cache_status)
            self.assertIn("AD cache saved:", calculator.last_cache_status)

    def test_concurrent_initialization_shares_one_successful_load(self):
        calculator = ADCalculator(n_neighbors=2)
        load_started = threading.Event()
        release_load = threading.Event()
        follower_waiting = threading.Event()
        calls = []
        results = []

        class TrackingCondition(threading.Condition):
            def wait(self, timeout=None):
                follower_waiting.set()
                return super().wait(timeout)

        calculator._load_condition = TrackingCondition()

        def load(_reference_path, cache_path=None, force_refit=False):
            calls.append((cache_path, force_refit))
            load_started.set()
            self.assertTrue(release_load.wait(timeout=2))
            calculator.fitted = True
            return calculator

        calculator.fit_from_excel_cached = load

        leader = threading.Thread(
            target=lambda: results.append(
                calculator.ensure_fitted_from_excel_cached("reference.xlsx")
            )
        )
        follower = threading.Thread(
            target=lambda: (
                results.append(calculator.ensure_fitted_from_excel_cached("reference.xlsx")),
            )
        )
        leader.start()
        self.assertTrue(load_started.wait(timeout=2))
        follower.start()
        self.assertTrue(follower_waiting.wait(timeout=2))
        release_load.set()
        leader.join(timeout=2)
        follower.join(timeout=2)

        self.assertFalse(leader.is_alive())
        self.assertFalse(follower.is_alive())
        self.assertEqual(len(calls), 1)
        self.assertEqual(results, [calculator, calculator])

    def test_concurrent_initialization_propagates_load_failure_to_waiter(self):
        calculator = ADCalculator(n_neighbors=2)
        load_started = threading.Event()
        release_load = threading.Event()
        follower_waiting = threading.Event()
        errors = []

        class TrackingCondition(threading.Condition):
            def wait(self, timeout=None):
                follower_waiting.set()
                return super().wait(timeout)

        calculator._load_condition = TrackingCondition()

        def load(_reference_path, cache_path=None, force_refit=False):
            load_started.set()
            self.assertTrue(release_load.wait(timeout=2))
            raise ValueError("reference is unreadable")

        calculator.fit_from_excel_cached = load

        def initialize():
            try:
                calculator.ensure_fitted_from_excel_cached("reference.xlsx")
            except Exception as exc:
                errors.append(str(exc))

        leader = threading.Thread(target=initialize)
        follower = threading.Thread(target=initialize)
        leader.start()
        self.assertTrue(load_started.wait(timeout=2))
        follower.start()
        self.assertTrue(follower_waiting.wait(timeout=2))
        release_load.set()
        leader.join(timeout=2)
        follower.join(timeout=2)

        self.assertFalse(leader.is_alive())
        self.assertFalse(follower.is_alive())
        self.assertEqual(len(errors), 2)
        self.assertTrue(any(error == "reference is unreadable" for error in errors))
        self.assertTrue(
            any(
                error == "AD reference load failed: reference is unreadable"
                for error in errors
            )
        )


if __name__ == "__main__":
    unittest.main()
