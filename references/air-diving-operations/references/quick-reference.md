# Quick Reference — Key Numbers, Limits, and Tables

## Environmental/Operational Limits

| Parameter | Limit |
|-----------|-------|
| Max air diving depth | 50 msw (165 fsw) |
| SRP max depth | 30 msw (exceptional 50 msw with risk assessment) |
| Max PPO2 (surface supplied) | 1.4 bar |
| Min O2 content | 200 mb PPO2 |
| Bail-out breathing rate | 50 LPM (40 IMCA base + 10 LPM safety factor) |
| Bail-out capacity rule | 1 min per 10m of umbilical from nearest safe haven |
| Ascent rate | 30 fsw/min (20 sec per 10 fsw) between stops |
| Surface decompress ascent | 40 fsw/min from 40 fsw to surface |
| Max compression rate | 100 fsw/min |
| Max surface interval (surface decomp) | 5 min |
| Max O2 in chamber | 23% |
| Chamber temp | Optimum 24°C (76°F) |
| Wind limit | 10 knots |
| Current limit (normal) | 1 knot |
| Current limit (with RA + monitoring) | 1-1.2 knots |
| Swell limit (DSV/stable platform) | 1.5 m |
| Swell limit (SRP/small craft) | 1.0 m |
| Wave: standard air decomp | Max 0.5-1 m |
| Wave: surface decomp | Max 1-2 m |
| Visibility limit | 500 m (standard) / 200 m (platform/moored) |
| Max in-water time per 24 hours | 4 hours |
| Standby diver minimum dives | 20 commercial dives |
| Standby diver Repetitive Group | A (no residual nitrogen) |
| DDC diameter | 60" minimum |
| DDC requirement | Two compartment (twin lock) |
| Oxygen reserve minimum | 90 m³ |
| Operational dive cease (O2) | < 90 m³ usable |
| Bend watch period | 18 hours after last dive |
| Post-dive near chamber | 4 hours within 20 min proximity, then 2 hours until 18 hrs post-surface |
| Pre-dive check frequency | Every 6 hours during continuous operations |
| Comms recording retention | 48 hours minimum |
| Max working hours | 12 hours continuous |
| Rest period | 12 hours unbroken |
| Longest shift | 24 hours before 8 hours rest |
| Diving team minimum | 1 DS, 2 divers, 2 DMs, 1 technician |
| SCUBA on worksite | NOT AUTHORIZED |
| Live boating | Class 2 or 3 DP DSV only |
| Repetitive diving | Emergency only, max 2 days |

## Pneumofathometer Correction
| Depth Range | Correction |
|-------------|------------|
| 0-100 fsw | +1 fsw |
| 101-200 fsw | +2 fsw |

## Bail-Out Calculation
```
Available Pressure (bar) = Charged Pressure - 20 bar (reserve)
Absolute Pressure at Depth (bar) = (depth_m / 10) + 1
Regulator allowance = 15 bar
Available for use (L) = Available_Pressure × 1000 ÷ (50 LPM × Absolute_Pressure)
```

Example: 38 msw, 10L bottle @ 170 bar
- Available = (170-20) × 1000 = 150,000 L
- At depth: 38/10 + 1 = 4.8 bar → 4.8 + 15 (reg) = 19.8 bar
- Rate: 50 × 4.8 = 240 L/min
- Duration: 15,000 / 240 = **6.25 min**

## O2 Toxicity — UPTD
```
UPTD = T × (2P - 1)^0.833
T = exposure time in minutes
P = oxygen partial pressure in atm abs
```

## Breathing Gas Quality (BS EN 12021)
| Contaminant | Max Limit |
|-------------|-----------|
| Carbon Monoxide | 5 ppm |
| Carbon Dioxide | 500 ppm |
| Oil | 0.5 mg/m³ |
| Water (compressor outlet) | 25 mg/m³ |
| Water (cylinder 40-200 bar) | 50 mg/m³ |
| Water (cylinder > 200 bar) | 35 mg/m³ |

## Compressor Log (Required Entries)
- Date, time started/finished pumping, accumulated hours
- Date filters changed, operator name & signature
- Entry in PMS

## Gas Log (Required Entries per quad/cylinder)
- Quad/cylinder number, initial pressure, initial O2 analysis
- Final pressure, final O2 analysis, contents at end of day
- Operator name & signature

## Purity Check Schedule
- On equipment acceptance
- After any significant repair/maintenance
- Every **6 monthly intervals**
- On acceptance and before going online
- Before each dive and during operations
- On-line analysers calibrated at **start of each shift**

## Checklists (Every 6 Hours During Continuous Operations)
1. Control van
2. Dive stages/baskets
3. Dive stage A-frame and tugger
4. Oxygen quads
5. Air quads
6. Breathing air compressors
7. Divers' helmets
8. DDC (subject to type)

## Per Dive Checks
- Pre-dive control van
- Pre-dive deck/chamber
- Pre-dive diver "call out" (each diver + standby)

## Post Dive Checks
- Post-dive supervisor's control van
- Post decompression DDC checks
- Post-dive equipment checks

## Line Signals — Royal Navy Protocol
- **Pull**: relatively long steady tension on line
- **Bells**: short tugs in pairs (or pairs + remaining odd bell)
- **All signals**: preceded by one pull to attract attention
- **Acknowledge**: repeat signal until correct response received
- **Only used when all comms failed**

## O2 Toxicity — CNS Symptoms (VENTID)
| Letter | Symptom |
|--------|---------|
| V | Vision (disturbances) |
| E | Ears (ringing) |
| N | Nausea |
| T | Twitching (muscles) |
| I | Irritability |
| D | Dizziness |

## Treatment Table Quick Guide
| Condition | Table |
|-----------|-------|
| Type I only, complete neuro exam normal | Table 5 |
| Severe Type I, any neurological symptom, Type II, AGE at 60 f | Table 6 (MANDATORY) |
| Severe symptoms unchanged/worsening at 60 f within 20 min | Table 6A |
| Deterioration at 60 f, needs deeper decomp | Table 4 |
| Non-responding severe AGE, life-threatening DCI | Table 7 |
| Deep blow-up > 60 min omitted decomp | Table 8 |

## Fitness to Return to Diving After DCI
| Condition | Lay-off |
|-----------|---------|
| Type I full recovery | 24 hours |
| Type I recurrence/relapse | 7 days |
| Type II full recovery + doctor approval | 7 days or longer |
| Type II residual | **Unfit for diving** |
| Pulmonary barotrauma/pneumothorax | 3 months + HRCT + specialist review |

## Emergency Drill Frequency
| Drill | Frequency |
|-------|-----------|
| Standby diver deployment | Weekly |
| Recovery of unconscious diver | Prior to 1st dive / weekly |
| Recovery of injured diver in water | 2 weekly |
| Loss of diver communications | 2 weekly |
| Loss/contamination of breathing gas | 2 weekly |
| Fouled or entrapped diver | 2 weekly |
| Fire or smoke hazard on deck | 2 weekly |
| LARS main winch / wire failure | 2 weekly |
| Onshore ER team comms | Monthly |

## IMCA/IOGP Reference Documents
- D014 — ICOP for Offshore Diving
- D016 — Underwater Air Lift Bags
- D018 — Examination & Testing of Diving Plant
- D023 — DESIGN SS Air Diving Systems
- D028 — Chain Lever Hoists
- D040 — DESIGN Mobile/Portable Systems
- D043 — Gas Cylinder Marking/Colour Coding
- D045 — Electricity Underwater
- D049 — HP Jetting Equipment
- D050 — Minimum Gas Quantities
- D054 — ROV During Diving Ops
- D064 — Cylinder/Valve Thread Compatibility
- D065 — Whip Checks
- D066 — Surface Swimming
- D067 — Currents on Divers Performance
- D070 — Use of Inert Gas
- D010 — Surface Umbilical Management
- D028 — Chain Lever Hoists
- M182 — Lifting Operations Guidelines
- OGP 376 — Lifting & Hoisting Safety
- OGP 471 — Oxy-Arc Underwater Cutting
- DMAC 003 — HP Water Jet Accidents
- DMAC 015 — Medical Equipment at Site
- DMAC 07 — Flying After Diving
- HSE DIS 5 — Exposure Limits for Air Diving
- BS EN 12021 — Breathing Gas Quality
- BS EN 1089-3:1997 — Cylinder Colour Coding
