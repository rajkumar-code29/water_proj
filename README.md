# Project SAGAR: cleaning Hussain Sagar, Hyderabad

**SAGAR = Solar Autonomous Grid for Aquatic Restoration.**
This is a concept design for cleaning polluted public lakes, starting with Hussain Sagar in Hyderabad. It combines:

1. **SAGAR-1 interceptor plant**: a 100 MLD Swiss/Swedish-style treatment plant built *across* the Kukatpally nala, the largest sewage inflow. The process train is MBR membranes (as at Stockholm's Henriksdal) plus ozone and activated carbon (Switzerland's "4th treatment stage"), then UV and aeration.
2. **Solar boat fleet**: 12 small 4.2 m solar catamarans on one platform with two payloads. **S-1 Skimmer** removes debris, hyacinth and algae scum. **O-1 Polisher** is the plant's process shrunk onto a boat: screen, filter, carbon, UV-C, oxygen.
3. **Sense & control**: a water-quality sonde on every boat, plus buoys, nala flow meters and a live public dashboard.

The deck is written for a **product design engineer**: the first half sets out the evidence, the second half is the engineering brief.

## What's in the repo

| Path | What it is |
|---|---|
| [`deck/Project_SAGAR_Hussain_Sagar.pptx`](deck/Project_SAGAR_Hussain_Sagar.pptx) | The presentation (29 slides, editable, with speaker notes) |
| [`deck/Project_SAGAR_Hussain_Sagar.pdf`](deck/Project_SAGAR_Hussain_Sagar.pdf) | PDF export of the same deck |
| [`research/hussain-sagar-pollution-data.md`](research/hussain-sagar-pollution-data.md) | Evidence pack: every pollution figure with its source and date, including where sources disagree |
| [`design/engineering-spec.md`](design/engineering-spec.md) | Engineering brief: boat platform spec, energy budget, BOM, plant process train, effluent targets, roadmap, KPIs |
| [`build/`](build/) | Node script that generates the deck (pptxgenjs). Edit the content there and rebuild |

## Key findings (as of November 2026)

- **At least ~40 million litres/day enters untreated.** Nala inflow is ~95 MLD on a study basis (2026 press reports say ~450 MLD) against 55 MLD of STP capacity.
- **Dissolved oxygen averages ~1.9 mg/L** (2016–20 data), below the 4 mg/L that fish need. It hit **zero** at three sites during the 2025 Ganesh immersion.
- **BOD is 24–76 mg/L**, 8–25× the bathing-water limit of 3 mg/L. Coliform reached 1,600 MPN/100 mL at 8 of 9 sampling points (2023).
- **September 2026 immersion (TGPCB):** turbidity peaked at 200 NTU, COD at 164 mg/L and TSS at 99 mg/L. The lake had not recovered by the post-immersion sampling.
- **The clue:** in the 2020 lockdown, BOD fell to ~8 mg/L. Cutting the inflow works.

TGPCB figures are as reported by the press. Request the original reports before any public or funding use.

## Rebuild the deck

```bash
cd build
npm install
npm run build          # writes ../deck/Project_SAGAR_Hussain_Sagar.pptx
```

To make the PDF, open the .pptx in PowerPoint and export to PDF, or run LibreOffice:

```bash
soffice --headless --convert-to pdf --outdir ../deck ../deck/Project_SAGAR_Hussain_Sagar.pptx
```

## Status

Concept v0.1 for discussion. Engineering numbers (energy, cost, footprint) are first-pass estimates, not certified designs.
