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

    def test_stable_below_limit(self):
        """Hops below the platform limit should be reported as stable."""
        hops_data = {'max_hops': 2}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'STABLE')
        self.assertEqual(verdict['color'], 'green')

    def test_at_limit(self):
        """Hops equal to the platform limit should be reported as at limit."""
        hops_data = {'max_hops': 4}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'AT LIMIT')
        self.assertEqual(verdict['color'], 'orange')
        self.assertIsNotNone(verdict['warning'])

    def test_apple_silicon_at_limit(self):
        """Apple Silicon has a lower limit, so it hits AT LIMIT earlier."""
        hops_data = {'max_hops': 3}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'mac_apple_silicon')
        self.assertEqual(verdict['status'], 'AT LIMIT')
        self.assertEqual(verdict['color'], 'orange')
        self.assertIsNotNone(verdict['warning'])

    def test_unstable_above_limit(self):
        """Hops above the platform limit should be reported as unstable."""
        hops_data = {'max_hops': 6}
        stability = self.analyzer.assess_stability(hops_data)
        verdict = self._verdict_for(stability, 'windows_x86')
        self.assertEqual(verdict['status'], 'UNSTABLE')
        self.assertEqual(verdict['color'], 'red')
        self.assertIsNotNone(verdict['warning'])

    def test_overall_worst_reflects_worst_status(self):
        """overall_worst should track the worst status among all platforms."""
        hops_data = {'max_hops': 6}
        stability = self.analyzer.assess_stability(hops_data)
        self.assertEqual(stability['overall_worst'], 'UNSTABLE')


if __name__ == '__main__':
    unittest.main()
