# 🔍 ProAV Shōko

**ProAV Shōko** – Din plattformsoberoende USB-detektiv för AV-miljöer.  
Analysera, verifiera och felsök USB-anslutningar i mötesrum och BYOD-miljöer.

[![Build Status](https://github.com/klangche/proav-shoko/actions/workflows/build.yml/badge.svg)](https://github.com/klangche/proav-shoko/actions)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/klangche/proav-shoko/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](https://github.com/klangche/proav-shoko)

---

## 📋 Innehållsförteckning

- [Översikt](#-översikt)
- [Varför ProAV Shōko?](#-varför-proav-shōko)
- [Funktioner](#-funktioner)
- [Flödesschema](#-flödesschema)
- [Installation](#-installation)
- [Användning](#-användning)
- [Teknisk stack](#-teknisk-stack)
- [Bygga från källa](#-bygga-från-källa)
- [Felsökning](#-felsökning)
- [Licens](#-licens)

---

## 🎯 Översikt

ProAV Shōko är ett **plattformsoberoende** analysverktyg som:

- 📡 **Skannar** alla anslutna USB-enheter
- 🌳 **Bygger** ett hierarkiskt träd över USB-kedjan
- 📊 **Beräknar** hops (antal nivåer) och tiers (djup)
- 🟢🟡🔴 **Bedömer** stabilitet baserat på kedjans längd
- 🖥️ **Visar** anslutna skärmar med upplösning
- 📄 **Genererar** professionella HTML- och PDF-rapporter

> **Perfekt för:** AV-tekniker, IT-support, säljare och diagnostikteam som behöver snabbt identifiera USB-problem i konferensrum.

---

## 🤔 Varför ProAV Shōko?

I moderna mötesrum ser vi ofta:

- USB-C-dockor (Unisynk, HP, Lenovo, CalDigit, Logitech, TiGHT, Hyper, Targus...)
- Flera hubbar i kedja
- Webbkameror, högtalartelefoner, pekpaneler, trådlösa presentationsdonglar, externa diskar
- Användarens egna iPads/iPhones/Android-enheter

**Problemet:** Långa kedjor orsakar ofta problem **endast på Apple Silicon Macs** (M1/M2/M3/M4), medan Windows och Intel Macs vanligtvis fungerar felfritt.

**ProAV Shōko** hjälper tekniker att bevisa:  
→ *"Kedjan har 5 hops → Windows & Intel OK, men Apple Silicon är inte stabilt"*

---

## ✨ Funktioner

| Funktion | Beskrivning | Status |
|----------|-------------|--------|
| 🌳 **USB-träd** | Hierarkisk vy över alla anslutna enheter | ✅ |
| 📏 **Hops & Tiers** | Beräkna antal nivåer och maximalt djup | ✅ |
| 🟢🟡🔴 **Stabilitetsbedömning** | Färgkodad baserad på kedjans längd | ✅ |
| 🍎 **Apple Silicon-varning** | Speciell varning vid 5+ hops | ✅ |
| 🖥️ **Skärminformation** | Visa anslutna skärmar med upplösning | ✅ |
| 📄 **HTML-rapport** | Mörk bakgrund, identisk med terminalen | ✅ |
| 📄 **PDF-rapport** | Lång, kontinuerlig sida för utskrift | ✅ |
| 🔄 **Automatisk öppning** | Rapporter öppnas direkt i webbläsare/PDF-visare | ✅ |
| ⚡ **Realtidsövervakning** | USB-anslutning/händelser (kräver admin) | ❌ *Ej portad* |
| 🔌 **PD-information** | Power Delivery-detaljer | ❌ *Ej portad* |

---

## 🔄 Flödesschema

```mermaid
graph TD
    Start([Start]) --> Platform[Plattformsinformation]
    Platform --> USB[Skanna USB-enheter]
    USB --> Tree[Bygg USB-träd]
    Tree --> Hops[Beräkna Hops & Tiers]
    Hops --> Stability{Stabilitetsbedömning}
    
    Stability --> CheckHops{Hops <= 3?}
    CheckHops -->|Ja| Green[🟢 STABIL]
    CheckHops -->|Nej| CheckHops5{Hops <= 5?}
    CheckHops5 -->|Ja| Yellow[🟡 OK]
    CheckHops5 -->|Nej| Red[🔴 INSTABIL]
    
    CheckHops5 --> AppleSilicon{Apple Silicon?}
    AppleSilicon -->|Ja & Hops=5| Orange[🟠 OSÄKER - Varning]
    AppleSilicon -->|Nej| Yellow
    
    Green --> Display[Skanna skärmar]
    Yellow --> Display
    Orange --> Display
    Red --> Display
    
    Display --> Report[Generera rapporter]
    Report --> HTML[HTML-rapport]
    Report --> PDF[PDF-rapport]
    
    HTML --> Open[Öppna rapporter]
    PDF --> Open
    Open --> Done([Klart])

    style Start fill:#1a1a2e,color:#fff
    style Done fill:#1a1a2e,color:#fff
    style Green fill:#00cc66,color:#000
    style Yellow fill:#ffcc00,color:#000
    style Orange fill:#ff8800,color:#000
    style Red fill:#ff3333,color:#fff
    style AppleSilicon fill:#ff8800,color:#000
