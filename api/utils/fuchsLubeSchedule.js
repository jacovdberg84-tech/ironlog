// Trusted site lubrication data transcribed from “Lube Schedule - Fuchs.xlsx”.
// Capacities are full-system capacities, not automatic top-up quantities.

export const FUCHS_LUBE_SOURCE = Object.freeze({
  name: "Fuchs equipment lube schedule",
  file_name: "Lube Schedule - Fuchs.xlsx",
  note: "Use the listed product and capacity for the matched model/component. A capacity is a full-system reference, not a top-up quantity.",
});

const PROFILES = [
  {
    "model": "CAT 140H",
    "category": "Grader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 29
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 38
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION & DIFFERENTIAL",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 47
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "TANDEM (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "CIRCLE DRIVE HOUSING",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 7
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FRONT WHEEL SPINDLE BEARING HOUSING",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 0.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 E2 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 140K",
    "category": "Grader",
    "components": [
      {
        "component": "ENGINE",
        "product": "Fuchs Titan Cargo MC 10W40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 18
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "Fuchs Renolin HO68 Hydraulic Oil",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 55
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION & DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 47
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "TANDEM (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "CIRCLE DRIVE HOUSING",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 7
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FRONT WHEEL SPINDLE BEARING HOUSING",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 0.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 40
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D5G",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "Fuchs Titan Cargo MC 10W40",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 13
          },
          {
            "label": "FDH",
            "litres": 11.5
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "Fuchs Renolin HO68 Hydraulic Oil",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 84
          },
          {
            "label": "FDH",
            "litres": 73
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 16
          },
          {
            "label": "FDH",
            "litres": 16
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 23
          },
          {
            "label": "FDH",
            "litres": 23
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D5K",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 13
          },
          {
            "label": "FDH",
            "litres": 11.5
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 84
          },
          {
            "label": "FDH",
            "litres": 73
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 16
          },
          {
            "label": "FDH",
            "litres": 16
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "WGB",
            "litres": 23
          },
          {
            "label": "FDH",
            "litres": 23
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D6R",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 28
          },
          {
            "label": "S6X",
            "litres": 28
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 47.3
          },
          {
            "label": "S6X",
            "litres": 51.5
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 148
          },
          {
            "label": "S6X",
            "litres": 146
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 13.5
          },
          {
            "label": "S6X",
            "litres": 13.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 5
          },
          {
            "label": "S6X",
            "litres": 1.9
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 25
          },
          {
            "label": "S6X",
            "litres": 25
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 70
          },
          {
            "label": "S6X",
            "litres": 77
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D6R2",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 28
          },
          {
            "label": "S6X",
            "litres": 28
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 47.3
          },
          {
            "label": "S6X",
            "litres": 51.5
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 148
          },
          {
            "label": "S6X",
            "litres": 146
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 13.5
          },
          {
            "label": "S6X",
            "litres": 13.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 5
          },
          {
            "label": "S6X",
            "litres": 1.9
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 25
          },
          {
            "label": "S6X",
            "litres": 25
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "BLT/TBC",
            "litres": 70
          },
          {
            "label": "S6X",
            "litres": 77
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D7R",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "AEC",
            "litres": 34
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "AEC",
            "litres": 54
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "AEC",
            "litres": 129
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "AEC",
            "litres": 13
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "AEC",
            "litres": 24
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "AEC",
            "litres": 57
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D8R",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 42
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 72
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 170
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 12
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 44
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 37
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 84
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D8T",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 42
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 72
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 170
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 12
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 44
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 37
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 84
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D8GC",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 42
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 72
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 170
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 12
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 44
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "TITAN UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 37
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 84
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LUBRENE LX 220 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D9R",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT D10T",
    "category": "Dozer",
    "components": [
      {
        "component": "ENGINE",
        "product": "Fuchs Titan Cargo MC 10W40",
        "capacities_l": [],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "Fuchs Titan Cargo MC 10W40",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "POWER TRAIN",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FINAL DRIVES (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "PIVOT SHAFT compartment",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "RECOIL SPRING COMPARTMENT",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 305CR",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 7
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 48
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 1
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 12
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 Ep2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 320D",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "ZCS",
            "litres": 22
          },
          {
            "label": "MDE",
            "litres": 30
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "ZCS",
            "litres": 138
          },
          {
            "label": "MDE",
            "litres": 138
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "ZCS",
            "litres": 8
          },
          {
            "label": "MDE",
            "litres": 8
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "ZCS",
            "litres": 8
          },
          {
            "label": "MDE",
            "litres": 8
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "ZCS",
            "litres": 26
          },
          {
            "label": "MDE",
            "litres": 26
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 Ep2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 329D",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 29.5
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 145
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 10
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 6
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 31
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 330D",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35.5
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 194
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 19
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 336D",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35.5
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 194
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 19
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LIM 450 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 340D",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35.5
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 194
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 19
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 349D",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 42
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 400
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 10
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 15
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 36
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 350",
    "category": "Excavator",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 32
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 10W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 217
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "SWING DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 15
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVE (each)",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 22
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 84
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 938H",
    "category": "Loader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50 + Caterpillar 1U-9891",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50 + Caterpillar 1U-9891",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 938K",
    "category": "Loader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "TITAN UTTO PRO",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "TITAN UTTO PRO",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 950GC",
    "category": "Loader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 20
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "Fuchs Titan UTTO TO-4 SAE 30",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 120
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 45
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 39
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 41
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 45.1
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT966H",
    "category": "Loader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 110
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 44
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50+ Caterpillar 1U-9891",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50+ Caterpillar 1U-9891",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 39
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT966L",
    "category": "Loader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 110
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 44
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50+ Caterpillar 1U-9891",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50+ Caterpillar 1U-9891",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 39
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT966GC",
    "category": "Loader",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 35
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULIC SYSTEM",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 110
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 44
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50+ Caterpillar 1U-9891",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50+ Caterpillar 1U-9891",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 64
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 39
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "BELL B25D",
    "category": "ADT",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO MC SAE 10W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 26
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "RENOLIN HO 68V",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 178
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "TITAN ATF 5668",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 34
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 45
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "MIDDLE DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 45
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 45
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "TRANSFER BOX",
        "product": "TITAN ATF 5668",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LMG 960 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LX-PEP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "BELL B30E",
    "category": "ADT",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO MC SAE 10W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 29
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "RENOLIN HO 68V",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 100
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "TITAN ATF 5668",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 30
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 54
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "MIDDLE DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 54
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 54
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "TRANSFER BOX (Gear Ratio)",
        "product": "TITAN ATF 5668",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "TRANSFER BOX (Kessler)",
        "product": "Dextron III (G/H) + additive",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN DP 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 65
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LMG 960 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LX-PEP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "BELL B40D",
    "category": "ADT",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO MC SAE 10W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 38
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "RENOLIN HO 68V",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 256
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "TITAN ATF 5668",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 34
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 60
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "MIDDLE DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 60
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 60
          }
        ],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "WET DISC BRAKES COOLING RESERVOIR",
        "product": "TITAN AFGRIFARM UTTO MP",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "TRANSFER BOX",
        "product": "TITAN ATF 5668",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LMG 960 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LX-PEP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "BELL B45E",
    "category": "ADT",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO MC SAE 10W-40",
        "capacities_l": [],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "RENOLIN HO 68V",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "TITAN ATF 5668",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "MIDDLE DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL & FINAL DRIVES",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [],
        "oil_drain_interval": 2000,
        "filter_change_interval": null
      },
      {
        "component": "WET DISC BRAKES COOLING RESERVOIR",
        "product": "Dextron III (G/H) + additive",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "TRANSFER BOX (Gear Ratio)",
        "product": "TITAN ATF 5668",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "TRANSFER BOX (Kessler)",
        "product": "Fuchs Titan Supergear 80W90 GL5 LS",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN DP 50",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "RENOLIT LUBRENE LMG 960 EP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "RENOLIT LX-PEP 2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 416D",
    "category": "TLB",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "L",
            "litres": 7.2
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "L",
            "litres": 49
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "L",
            "litres": 20.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "L",
            "litres": 11
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "L",
            "litres": 24
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL (Steerable axle)",
        "product": "Fuchs Titan UTTO TO-4 SAE 30",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVES",
        "product": "TITAN SUPERGEAR SAE 80W-90",
        "capacities_l": [
          {
            "label": "L",
            "litres": 0.7
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "L",
            "litres": 25.5
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 416E",
    "category": "TLB",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 7.2
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "TITAN TRUCK PLUS SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 40
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 18.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 11
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 24
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 0.7
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 25.5
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 426",
    "category": "TLB",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8.8
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "TITAN UTTO TO-4 SAE 10W",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 40
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 15
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 11
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 16
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVES",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 1.4
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 22
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "CAT 428F",
    "category": "TLB",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO SAE 15W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 7.6
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "HYDRAULICS",
        "product": "TITAN UTTO TO-4 SAE 10W",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 42
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": 500
      },
      {
        "component": "TRANSMISSION",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 19
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 500
      },
      {
        "component": "FRONT DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 11
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "REAR DIFFERENTIAL",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 16.5
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVES REAR",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 1.7
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "FINAL DRIVES FRONT",
        "product": "Fuchs Titan UTTO TO-4 SAE 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 0.7
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "Fuchs Maintain Fricofin Longlife Premix",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 19.5
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "ALL MERCEDES BENZ TRUCKS",
    "category": "Truck",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO MC SAE 10W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 39
          }
        ],
        "oil_drain_interval": 500,
        "filter_change_interval": 250
      },
      {
        "component": "GEARBOX",
        "product": "TITAN SUPERGEAR MC SAE 80W90",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 16
          }
        ],
        "oil_drain_interval": 1000,
        "filter_change_interval": 1000
      },
      {
        "component": "DIFFERENTIAL & FINAL DRIVES",
        "product": "TITAN SUPERGEAR MC SAE 80W90",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "STEERING",
        "product": "TITAN ATF 5500",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "HYDRAULICS",
        "product": "RENOLIN HO 68 V",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "BRAKE AND CLUTCH",
        "product": "MAINTAIN DOT 4",
        "capacities_l": [],
        "oil_drain_interval": 1000,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 40
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Pins and Bushes",
        "product": "Fuchs Renolit Lubrene LiM450 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "GREASE - Bearings and Propshafts",
        "product": "Fuchs Renolit Lubrene LX220 EP2",
        "capacities_l": [],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  },
  {
    "model": "TOYOTA HILUX 2.4 GD6 4x4",
    "category": "LDV",
    "components": [
      {
        "component": "ENGINE",
        "product": "TITAN CARGO MC SAE 10W-40",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8
          }
        ],
        "oil_drain_interval": 10000,
        "filter_change_interval": 10000
      },
      {
        "component": "GEARBOX",
        "product": "TITAN SUPERGEAR MC SAE 80W90",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 10
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "DIFFERENTIAL Front",
        "product": "TITAN SUPERGEAR MC SAE 80W90",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "DIFFERENTIAL Rear",
        "product": "TITAN SUPERGEAR MC SAE 80W90",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 8
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      },
      {
        "component": "COOLING SYSTEM",
        "product": "MAINTAIN FRICOFIN LL 50",
        "capacities_l": [
          {
            "label": "CAPACITY",
            "litres": 20
          }
        ],
        "oil_drain_interval": null,
        "filter_change_interval": null
      }
    ]
  }
];

function normalize(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function cloneComponents(components = []) {
  return components.map((component) => ({
    component: String(component.component || ""),
    product: String(component.product || ""),
    capacities_l: Array.isArray(component.capacities_l)
      ? component.capacities_l.map((capacity) => ({
          label: String(capacity.label || "CAPACITY"),
          litres: Number(capacity.litres || 0),
        })).filter((capacity) => Number.isFinite(capacity.litres) && capacity.litres > 0)
      : [],
    oil_drain_interval: Number(component.oil_drain_interval || 0) || null,
    filter_change_interval: Number(component.filter_change_interval || 0) || null,
  }));
}

function matchesModel(profile, equipmentName, category) {
  const model = normalize(profile?.model);
  const equipment = normalize(equipmentName);
  const assetCategory = normalize(category);
  if (!model) return false;

  if (model === "ALLMERCEDESBENZTRUCKS") {
    return /MERCEDES|BENZ|ACTROS|AXOR/.test(equipment)
      || (/TRUCK|TIPPER/.test(assetCategory) && /MERCEDES|BENZ|ACTROS|AXOR/.test(equipment));
  }

  return Boolean(equipment && equipment.includes(model));
}

export function getFuchsLubricationProfile({ equipmentName = "", category = "" } = {}) {
  const matched = [...PROFILES]
    .sort((a, b) => normalize(b.model).length - normalize(a.model).length)
    .find((profile) => matchesModel(profile, equipmentName, category));

  if (!matched) return null;
  return {
    source: FUCHS_LUBE_SOURCE.name,
    matched_model: String(matched.model || ""),
    equipment_category: String(matched.category || ""),
    components: cloneComponents(matched.components),
  };
}

function uniqueCapacity(component) {
  const capacities = Array.isArray(component?.capacities_l) ? component.capacities_l : [];
  if (!capacities.length) return null;
  const litres = [...new Set(capacities.map((capacity) => Number(capacity.litres || 0)).filter((value) => value > 0))];
  return litres.length === 1 ? litres[0] : null;
}

function isDueAtService(component, serviceHours) {
  const interval = Number(component?.oil_drain_interval || 0);
  const service = Number(serviceHours || 0);
  return interval > 0 && service > 0 && service % interval === 0;
}

export function getFuchsServiceRequirements(profile, serviceHours) {
  if (!profile?.components?.length) return [];
  return profile.components
    .filter((component) => isDueAtService(component, serviceHours))
    .map((component) => ({
      component: String(component.component || ""),
      product: String(component.product || ""),
      system_capacity_l: uniqueCapacity(component),
      capacity_options_l: (component.capacities_l || []).map((capacity) => ({
        label: String(capacity.label || "CAPACITY"),
        litres: Number(capacity.litres || 0),
      })),
      oil_drain_interval: Number(component.oil_drain_interval || 0) || null,
      filter_change_interval: Number(component.filter_change_interval || 0) || null,
    }));
}

export function buildFuchsServiceOilItems(profile, serviceHours) {
  return getFuchsServiceRequirements(profile, serviceHours)
    .filter((item) => Number(item.system_capacity_l || 0) > 0)
    .map((item) => ({
      name: `${item.component}: ${item.product}`,
      qty: Number(item.system_capacity_l),
      unit: "L",
      part_hint: String(item.product || ""),
      component: String(item.component || ""),
      source: FUCHS_LUBE_SOURCE.name,
    }));
}

export function formatFuchsCapacity(component) {
  const capacities = Array.isArray(component?.capacities_l) ? component.capacities_l : [];
  if (!capacities.length) return "Capacity not listed";
  const labelled = capacities.map((capacity) => {
    const label = String(capacity.label || "").trim().toUpperCase();
    const litres = Number(capacity.litres || 0);
    const prefix = label && label !== "CAPACITY" && label !== "L" ? `${label}: ` : "";
    return `${prefix}${litres} L`;
  });
  return labelled.join(" / ");
}

