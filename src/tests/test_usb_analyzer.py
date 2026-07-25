"""
Unit tests for the USB analyzer
"""

import unittest
from src.usb_analyzer import USBAnalyzer


class TestUSBAnalyzer(unittest.TestCase):
    """Test cases for USBAnalyzer."""

    def setUp(self):
        self.analyzer = USBAnalyzer()

    def _verdict_for(self, stability, platform_id):
        return next(v for v in stability['verdicts'] if v['id'] == platform_id)

    def test_stable_below_both_limits(self):
        """Hops and tiers both below the platform limit should be stable."""
        hops_data = {'max_hops': 2, 'max_tiers': 2}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'STABLE')
        self.assertEqual(verdict['color'], 'green')
        self.assertIsNone(verdict['warning'])

    def test_hops_at_limit_tiers_stable(self):
        """If hops is at the limit but tiers is stable, overall is AT LIMIT."""
        hops_data = {'max_hops': 4, 'max_tiers': 2}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'AT LIMIT')
        self.assertEqual(verdict['color'], 'orange')
        self.assertIn('Hops', verdict['warning'])
        self.assertNotIn('Tiers', verdict['warning'])

    def test_tiers_unstable_overrides_hops_stable(self):
        """If tiers exceeds its limit while hops is fine, overall is UNSTABLE."""
        hops_data = {'max_hops': 1, 'max_tiers': 6}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'UNSTABLE')
        self.assertEqual(verdict['color'], 'red')
        self.assertIn('Tiers', verdict['warning'])
        self.assertNotIn('Hops', verdict['warning'])

    def test_both_unstable_reports_both_warnings(self):
        """If both dimensions exceed their limits, both show up in the warning."""
        hops_data = {'max_hops': 6, 'max_tiers': 6}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'UNSTABLE')
        self.assertIn('Hops', verdict['warning'])
        self.assertIn('Tiers', verdict['warning'])

    def test_apple_silicon_lower_limit(self):
        """Apple Silicon has a lower limit, so it hits AT LIMIT earlier than Windows."""
        hops_data = {'max_hops': 3, 'max_tiers': 1}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'mac_apple_silicon')
        self.assertEqual(verdict['status'], 'AT LIMIT')

    def test_overall_worst_reflects_worst_status(self):
        """overall_worst should track the worst status among all platforms."""
        hops_data = {'max_hops': 6, 'max_tiers': 1}
        stability = self.analyzer.assess_stability(hops_data)
        self.assertEqual(stability['overall_worst'], 'UNSTABLE')

    def test_save_and_reload_hop_limits_csv(self):
        """save_hop_limits_csv should write a CSV that _load_limits can read back."""
        import tempfile, os
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as tmp:
            tmp_path = tmp.name
        try:
            self.analyzer.save_hop_limits_csv(tmp_path)
            reloaded = USBAnalyzer(tmp_path)
            self.assertEqual(reloaded.hop_limits.get('windows_x86'), self.analyzer.hop_limits.get('windows_x86'))
            self.assertEqual(reloaded.tier_limits.get('windows_x86'), self.analyzer.tier_limits.get('windows_x86'))
        finally:
            os.unlink(tmp_path)


if __name__ == '__main__':
    unittest.main()
