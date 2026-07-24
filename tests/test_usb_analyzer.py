"""
Enhetstester för USB-analysatorn
"""

import unittest
from src.usb_analyzer import USBAnalyzer


class TestUSBAnalyzer(unittest.TestCase):
    """Testfall för USBAnalyzer."""

    def test_assess_stability(self):
        """Testa stabilitetsbedömning."""
        analyzer = USBAnalyzer()

        # Hops <= 3
        hops_data = {'max_hops': 2}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'STABIL')
        self.assertEqual(result['color'], 'green')

        # Hops = 4
        hops_data = {'max_hops': 4}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'OK')
        self.assertEqual(result['color'], 'yellow')

        # Hops = 5 Apple Silicon
        hops_data = {'max_hops': 5}
        result = analyzer.assess_stability(hops_data, True)
        self.assertEqual(result['status'], 'OSÄKER')
        self.assertEqual(result['color'], 'orange')
        self.assertIsNotNone(result['warning'])

        # Hops = 5 Intel
        hops_data = {'max_hops': 5}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'OK')
        self.assertEqual(result['color'], 'yellow')

        # Hops >= 6
        hops_data = {'max_hops': 6}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'INSTABIL')
        self.assertEqual(result['color'], 'red')


if __name__ == '__main__':
    unittest.main()
