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

        # Testa hops <= 3
        hops_data = {'max_hops': 2}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'STABIL')

        # Testa hops = 4
        hops_data = {'max_hops': 4}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'OK')

        # Testa hops = 5 Apple Silicon
        hops_data = {'max_hops': 5}
        result = analyzer.assess_stability(hops_data, True)
        self.assertEqual(result['status'], 'OSÄKER')
        self.assertIsNotNone(result['warning'])

        # Testa hops >= 6
        hops_data = {'max_hops': 6}
        result = analyzer.assess_stability(hops_data, False)
        self.assertEqual(result['status'], 'INSTABIL')


if __name__ == '__main__':
    unittest.main()
