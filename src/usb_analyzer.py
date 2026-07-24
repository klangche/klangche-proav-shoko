"""
USB-analysator - hanterar USB-träd, hops, tiers och stabilitetsbedömning
"""

import sys
from collections import defaultdict
from typing import Dict, List, Any, Optional

try:
    from usbmonitor import USBMonitor
    from usbmonitor.attributes import ID_MODEL, ID_VENDOR, DEVNAME, ID_MODEL_ID, ID_VENDOR_ID
except ImportError:
    print("❌ usbmonitor är inte installerat. Kör: pip install usbmonitor")
    sys.exit(1)


class USBAnalyzer:
    """Analyserar USB-enheter och bygger hierarkiskt träd."""

    def __init__(self):
        """Initiera USB-analysatorn."""
        self.monitor = USBMonitor()
        self.devices = {}

    def build_tree(self) -> List[Dict[str, Any]]:
        """
        Bygger ett hierarkiskt träd av USB-enheter.

        Returns:
            Lista med trädstruktur för USB-enheter.
        """
        # Hämta alla USB-enheter
        try:
            # Försök att använda devpath för att bygga träd
            dev_dict = self.monitor.get_available_devices()
            self.devices = dev_dict

            # Konvertera till trädstruktur
            tree = []
            # Gruppera enheter efter devpath
            for dev_name, attributes in dev_dict.items():
                # Extrahera devpath
                devpath = attributes.get('devpath', '')

                # Skapa nod
                node = {
                    'name': dev_name,
                    'devpath': devpath,
                    'attributes': attributes,
                    'children': [],
                    'is_hub': self._is_hub(attributes),
                    'model': attributes.get(ID_MODEL, 'Okänd modell'),
                    'vendor': attributes.get(ID_VENDOR, 'Okänd tillverkare'),
                    'product': attributes.get('ID_MODEL', 'Okänd produkt')
                }

                # Hitta rätt plats i trädet
                if not devpath or devpath.count('/') <= 1:
                    # Rotnivå
                    tree.append(node)
                else:
                    # Försök hitta förälder
                    parent_path = '/'.join(devpath.split('/')[:-1])
                    parent = self._find_node_by_path(tree, parent_path)
                    if parent:
                        parent['children'].append(node)
                    else:
                        # Om förälder inte finns, lägg på rotnivå
                        tree.append(node)

            return tree

        except Exception as e:
            print(f"⚠️  Kunde inte läsa USB-enheter: {e}")
            return []

    def _is_hub(self, attributes: Dict) -> bool:
        """Kontrollera om enhet är en hub."""
        # Enkel heuristik: hubbar har ofta 'hub' i modellnamnet eller produkt-ID
        model = attributes.get(ID_MODEL, '').lower()
        product = attributes.get('ID_MODEL', '').lower()
        return 'hub' in model or 'hub' in product

    def _find_node_by_path(self, tree: List[Dict], path: str) -> Optional[Dict]:
        """Hitta nod med specifik devpath i trädet."""
        for node in tree:
            if node['devpath'] == path:
                return node
            # Sök i barn
            if node['children']:
                found = self._find_node_by_path(node['children'], path)
                if found:
                    return found
        return None

    def calculate_hops_and_tiers(self, tree: List[Dict]) -> Dict[str, Any]:
        """
        Beräknar hops och tiers från USB-trädet.

        Args:
            tree: USB-träd som lista med noder.

        Returns:
            Dictionary med hops, tiers och detaljer.
        """
        if not tree:
            return {
                'max_hops': 0,
                'max_tiers': 0,
                'devices_by_hops': {},
                'all_hops': []
            }

        # Beräkna djup för varje nod
        depths = []
        devices_by_hops = defaultdict(list)

        def traverse(node, depth):
            """Traversera trädet och samla djup."""
            # Använd devpath för att räkna hops (antal nivåer)
            if node['devpath']:
                hops = node['devpath'].count('/')
            else:
                hops = depth

            devices_by_hops[hops].append(node['name'])
            depths.append(hops)

            # Traversera barn
            for child in node['children']:
                traverse(child, depth + 1)

        # Börja traversering
        for root in tree:
            traverse(root, 0)

        max_hops = max(depths) if depths else 0
        max_tiers = len(set(depths)) if depths else 0

        return {
            'max_hops': max_hops,
            'max_tiers': max_tiers,
            'devices_by_hops': dict(devices_by_hops),
            'all_hops': depths
        }

    def assess_stability(self, hops_data: Dict[str, Any], is_apple_silicon: bool = False) -> Dict[str, Any]:
        """
        Bedömer stabilitet baserat på hops och plattform.

        Args:
            hops_data: Data från calculate_hops_and_tiers.
            is_apple_silicon: True om Apple Silicon.

        Returns:
            Dictionary med stabilitetsbedömning.
        """
        max_hops = hops_data.get('max_hops', 0)

        # Standardbedömning
        if max_hops <= 3:
            verdict = {
                'status': 'STABIL',
                'color': 'green',
                'label': '🟢 STABIL',
                'warning': None
            }
        elif max_hops <= 5:
            # Varning för Apple Silicon vid 5 hops
            if is_apple_silicon and max_hops == 5:
                verdict = {
                    'status': 'OSÄKER',
                    'color': 'orange',
                    'label': '🟠 OSÄKER',
                    'warning': '⚠️ Apple Silicon kan uppleva instabilitet vid 5 hops!'
                }
            else:
                verdict = {
                    'status': 'OK',
                    'color': 'yellow',
                    'label': '🟡 OK',
                    'warning': None
                }
        else:  # >= 6
            verdict = {
                'status': 'INSTABIL',
                'color': 'red',
                'label': '🔴 INSTABIL',
                'warning': '⚠️ Lång USB-kedja kan orsaka problem!'
            }

        # Extra varning för Apple Silicon
        if is_apple_silicon and max_hops >= 4:
            verdict['warning'] = (verdict['warning'] or '') + ' Apple Silicon rekommenderar max 4 hops.'

        return verdict
