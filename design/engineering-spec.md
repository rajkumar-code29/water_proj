# Project SAGAR: engineering design brief

**SAGAR = Solar Autonomous Grid for Aquatic Restoration.**
Audience: product design engineer. Status: concept, v0.1, 8 Nov 2026. All figures are first-pass engineering estimates for discussion. They are not certified designs.

The system has three parts. Each covers what the others can't:

| Part | Job | Why it exists |
|---|---|---|
| **SAGAR-1 interceptor plant** (across the Kukatpally nala) | Stop the source and treat the volume | ≥ 40 MLD of untreated water enters daily. Only a plant at the inlet can handle that volume |
| **Solar boat fleet** (S-1 Skimmer, O-1 Polisher) | Clean the surface and treat hotspots | The lake is nearly still. Fixed barriers need current, so the surface needs mobile units |
| **Sense & control** (sonde on every boat, 6 buoys, dashboard) | Measure everything and direct the boats | No current verified flow or quality time-series exists |

---

## 1. Solar boat platform

One platform, two payloads. The hull, power, propulsion and autonomy are shared, and the payload bay takes one of two modules.

### 1.1 Platform specification

| Item | Spec | Rationale |
|---|---|---|
| Hull | Twin-hull catamaran, rotomoulded **HDPE** pontoons, 316 SS / marine-aluminium frame | No corrosion at EC ~1,700 µS/cm. UV-stable, recyclable, cheap tooling at volume |
| Dimensions | LOA **4.2 m**, beam **2.2 m** (road-trailerable), draft **≤ 0.35 m** loaded | Shallow, silted margins. Fits a standard trailer |
| Mass | ~380 kg lightship; payload up to 150 kg | |
| Flotation | Foam-filled pontoon compartments; stays afloat with one pontoon breached | Safety in a public lake |
| Solar | **1.62 kWp** canopy (6 × 270 Wp mono-PERC, ~7.2 m²), MPPT 48 V | Also shades electronics and the debris bin |
| Battery | **9.6 kWh LiFePO4**, 48 V, as **4 × 2.4 kWh hot-swap packs (~22 kg each)** | LiFePO4 for thermal safety at 45 °C. One person can lift a 22 kg pack |
| Propulsion | 2 × 1 kW **rim-drive** electric thrusters, differential steering, no rudder | No exposed blades, so they're safe for swimmers and birds and resist hyacinth fouling |
| Speed | 1 kn skimming, 2 kn transit, 3 kn max (software-limited) | |
| Autonomy | RTK-GNSS, 2 × stereo cameras (debris detection), mmWave radar, edge compute | Radar works through glare and haze where LiDAR struggles |
| Comms | LTE primary + LoRa fallback, 1 Hz telemetry | |
| Safety | Remote and on-deck e-stop, geofence, collision avoidance (tour boats, paddle boats), nav lights, high-vis colour | Busy tourist lake |
| Sensing (every boat) | Multi-parameter sonde: DO, pH, EC, turbidity, temperature, chlorophyll-a | Each boat is also a mobile monitoring station |
| Ingress | IP67 enclosures, sealed connectors, passive-cooled battery bay | Monsoon |

### 1.2 Payload modules

**S-1 Skimmer (debris, hyacinth, algae scum)**
- Inclined mesh conveyor between the hulls. Intake width **1.6 m**, 5 mm mesh.
- **250 L** removable bin (~120 kg wet per load). Docks 3–4 times a day, so about **400 kg/day** maximum.
- Coverage: 1.6 m × 0.51 m/s ≈ 2,950 m²/h, so ~1.3 ha in a 4.5 h skim shift.

**O-1 Polisher (a 4-stage treatment plant shrunk onto a boat)**
Same process logic as the big plant, scaled down:
1. 2 mm intake screen
2. 25 µm self-cleaning disc filter
3. Activated-carbon cartridge (swappable; micropollutant demo, Swiss-style 4th stage)
4. UV-C LED reactor (~40 mJ/cm²)
5. Venturi oxygen injection, then discharge at 1.5 m depth

- Flow **10 m³/h**, about **60 m³/day**. Sized for immersion bays, algae hotspots, demos and education. **It does not treat the whole lake.** See the scale-math slide in the deck.

### 1.3 Daily energy budget (Hyderabad, annual average)

Solar yield = 1.62 kWp × 5.0 peak-sun-hours × 0.80 performance ratio ≈ **6.5 kWh/day**.
Monsoon days (~3.8 PSH) give ~4.9 kWh/day. Shorten payload hours or swap packs at the dock.

| Load (kWh/day) | S-1 Skimmer | O-1 Polisher |
|---|---|---|
| Propulsion | 3.9 (skim 0.45 kW × 6 h + transit 0.8 kW × 1.5 h) | 1.4 (station-keep 0.1 kW × 6 h + transit 0.8 kW × 1 h) |
| Payload | 0.9 (conveyor 0.15 kW × 6 h) | 2.5 (pump 0.30 + UV 0.12 kW × 6 h) |
| Electronics & sensing | 1.0 | 0.9 |
| **Total** | **5.8** | **4.8** |
| Solar yield | 6.5 | 6.5 |
| Battery reserve (7.7 kWh usable at 80% DoD) | ~1.3 days with no sun | ~1.6 days with no sun |

### 1.4 Indicative bill of materials (prototype, INR lakh)

| Subsystem | ₹ lakh |
|---|---|
| Hulls & frame | 3.0 |
| Solar canopy + MPPT | 1.2 |
| LiFePO4 packs + BMS | 3.0 |
| Propulsion (2 rim-drives + controllers) | 2.0 |
| Payload module (conveyor or filter/UV train) | 2.5 |
| Water-quality sonde | 2.5 |
| Autonomy (RTK, cameras, radar, compute) | 2.3 |
| Comms & safety | 0.5 |
| Assembly, test, contingency | 1.5 |
| **Total (prototype)** | **18.5** |
| Target at 50+ units | ≤ 12 |

### 1.5 Fleet

- **Phase 2 fleet: 12 boats** (8 × S-1, 4 × O-1) and **2 solar docks** (Lumbini Park, Sanjeevaiah Park). Each dock has a ~10 kWp canopy, pack-swap racks and a debris compactor.
- Debris capacity: 8 × 400 kg ≈ **3.2 t/day**. In immersion season the fleet runs 2 shifts on swapped packs. Heavy idol remains still need shore cranes; the boats handle the floating fraction.
- Coverage: 8 skimmers × 1.3 ha ≈ 10 ha/day. Debris collects in the wind-driven margins, nala mouths and immersion bays (~60 ha), so hotspots get swept about every 6 days. Camera detection sends boats where the debris is.

### 1.6 Design requirements

- **Environment:** air 10–46 °C; water 20–35 °C; UV index 11+; monsoon gusts ~60 km/h; pH 7–9.5; biofouling; algae scum; hyacinth; plastics; cloth; flowers; Plaster-of-Paris idol fragments.
- **Materials:** no mild steel; HDPE, 316 SS and anodised aluminium only; antifouling-free hull (foul-release coating or none).
- **Serviceability:** tool-less module swap in < 15 min; standard dock interface for mechanical, power and data; battery swap < 2 min per pack.
- **Safety:** max 3 kn; geofence around swimming and boating zones; e-stop reachable from both sides; fail-safe drift-and-anchor on comms loss.
- **Manufacture:** rotomoulding for the hulls (low tooling cost in India); off-the-shelf 48 V EV parts; target local content ≥ 70%.

### 1.7 Open questions for the product designer

1. HDPE rotomoulded vs. welded aluminium hulls: cost, repairability, mass.
2. Conveyor vs. suction intake for PoP slurry and algae scum.
3. Radar + camera vs. adding LiDAR: cost vs. reliability in haze and glare.
4. Automatic dock charging (contacts or inductive) vs. manual pack swap.
5. Anti-theft and vandalism: tracking, tamper alarms, lockable bays.
6. Can S-1 and O-1 share one hull with a 15-minute module swap, or should they be two dedicated SKUs?

---

## 2. SAGAR-1 interceptor plant

The Swiss and Swedish idea applied to Hussain Sagar: **treat the water where it enters, before it reaches the open lake.** The plant is built across the Kukatpally nala, the largest and most industrial inflow (70–155 MLD reported).

### 2.1 Process train

| # | Stage | Function |
|---|---|---|
| 1 | **Inlet weir across the nala** + 25 mm coarse screens, trash boom | Diverts all dry-weather flow into the plant. Storm surplus overflows the weir after screening |
| 2 | First-flush storage tank | Captures the dirtiest early-monsoon runoff for later treatment |
| 3 | 6 mm → 2 mm fine screens, vortex grit & grease removal | Protects the membranes |
| 4 | Biological nutrient removal (anoxic/aerobic + bio-P) | Removes BOD, nitrogen, phosphorus |
| 5 | **MBR ultrafiltration (0.04 µm)**, Henriksdal-style | Replaces clarifiers. Removes TSS and bacteria. Small footprint |
| 6 | **Ozonation + GAC filter**, Swiss "4th stage" | Removes ≥ 80% of micropollutants (pharmaceuticals, dyes, industrial organics) |
| 7 | UV disinfection | Final pathogen barrier |
| 8 | Cascade aeration outlet | Returns water to the lake at DO ≥ 6 mg/L |
| Side | Sludge → anaerobic digester → biogas CHP; struvite P-recovery; screenings → recycling/RDF | Energy and resource recovery |

### 2.2 Key numbers

| Item | Value |
|---|---|
| Design flow | **100 MLD average, 150 MLD peak**, as 4 × 25 MLD modular trains (covers the 70–155 MLD reported range; trains can be added) |
| Footprint | ~3.5–4 ha, covered or semi-sunken, with a **public park on the roof**. Hyderabad sits on Deccan granite, like Stockholm |
| Energy | ~0.9 kWh/m³, so ~90 MWh/day |
| Energy supply | Biogas CHP ~25–30% + up to **4 MWp** rooftop/canopy solar (~16 MWh/day, ~18%) + green-tariff grid for the rest |
| Indicative capex | **₹450–550 crore**. Benchmark: HMWSSB's 17-STP package, 376.5 MLD for ₹1,280.87 cr ≈ ₹3.4 cr/MLD, plus an MBR and 4th-stage premium |

### 2.3 Effluent targets vs. NGT (2019) discharge norms

| Parameter | NGT norm (metro) | SAGAR-1 target |
|---|---|---|
| BOD | ≤ 10 mg/L | **≤ 5** |
| COD | ≤ 50 mg/L | **≤ 30** |
| TSS | ≤ 20 mg/L | **≤ 2** |
| Total N | ≤ 10 mg/L | **≤ 8** |
| Total P | ≤ 1 mg/L | **≤ 0.3** |
| Faecal coliform | < 230 MPN/100 mL (desirable < 100) | **< 10** |
| Micropollutants | not regulated | **≥ 80% removal** |

---

## 3. Roadmap and KPIs

| Phase | Months | Scope |
|---|---|---|
| 0 Baseline | 0–3 | Flow meters on all 4 nalas (settles the 95 vs. 450 MLD question), bathymetry, sediment cores, 6 sensor buoys, permits (HMDA, TGPCB, GHMC) |
| 1 Prototype | 3–12 | 2 × S-1 + 1 × O-1. Pilot in the NTR Park / Lumbini Park bays. Field test in the Ganesh immersion, Sept 2027 |
| 2 Fleet + plant design | 12–30 | 12 boats + 2 solar docks. SAGAR-1 detailed design, DBO/HAM tender, construction start |
| 3 Plant + replicate | 30–48 | SAGAR-1 commissioning. Interceptors on Picket, Banjara, Balkapur. Fleet to other Hyderabad lakes |

| KPI | Baseline | Target (end of Phase 3) |
|---|---|---|
| Untreated dry-weather inflow | ≥ 40 MLD | **0** |
| Dissolved oxygen | ~0–3 mg/L | **≥ 4 mg/L** (CPCB Class D) |
| BOD | 24–76 mg/L | **≤ 10 mg/L**, then ≤ 3 (bathing) long-term |
| Total coliform | 1,600 MPN/100 mL | **< 500** (bathing) |
| Floating debris | — | **−90%** (camera survey index) |
| Public data | none | **Live 24×7 dashboard** |
