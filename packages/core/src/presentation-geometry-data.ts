// Generated standard data, not source document content. See the generator and license.
export interface PresetDefinition {
  readonly adjustments: readonly (readonly [string, string])[];
  readonly guides: readonly (readonly [string, string])[];
  readonly text: readonly [string, string, string, string] | null;
  readonly connections: readonly (readonly [string, string, string])[];
  readonly paths: readonly { readonly fill: boolean; readonly stroke: boolean; readonly commands: readonly (readonly string[])[] }[];
}
/*! @license
 * Selected geometric definition data from ECMA-376 Part 1, fifth edition (2016),
 * OfficeOpenXML-DrawingMLGeometries.zip / presetShapeDefinitions.xml.
 * Source: https://www.ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip
 * This project implements selected standard definitions; it does not modify or
 * republish the standard. The evaluator and rendering implementation are first-party.
 * 
 * COPYRIGHT NOTICE
 * © 2016 Ecma International
 * 
 * This document may be copied, published and distributed to others, and certain
 * derivative works of it may be prepared, copied, published, and distributed, in
 * whole or in part, provided that the above copyright notice and this copyright
 * license and disclaimer are included on all such copies and derivative works.
 * The only derivative works that are permissible under this copyright license and
 * disclaimer are:
 * 
 * (i) works which incorporate all or portion of this document for the purpose of
 * providing commentary or explanation (such as an annotated version of the document),
 * (ii) works which incorporate all or portion of this document for the purpose of
 * incorporating features that provide accessibility,
 * (iii) translations of this document into languages other than English and into
 * different formats and
 * (iv) works by making use of this specification in standard conformant products
 * by implementing (e.g. by copy and paste wholly or partly) the functionality therein.
 * 
 * However, the content of this document itself may not be modified in any way,
 * including by removing the copyright notice or references to Ecma International,
 * except as required to translate it into languages other than English or into a
 * different format.
 * 
 * The official version of an Ecma International document is the English language
 * version on the Ecma International website. In the event of discrepancies between
 * a translated version and the official version, the official version shall govern.
 * 
 * The limited permissions granted above are perpetual and will not be revoked by
 * Ecma International or its successors or assigns.
 * 
 * This document and the information contained herein is provided on an “AS IS” basis
 * and ECMA INTERNATIONAL DISCLAIMS ALL WARRANTIES, EXPRESS OR IMPLIED, INCLUDING
 * BUT NOT LIMITED TO ANY WARRANTY THAT THE USE OF THE INFORMATION HEREIN WILL NOT
 * INFRINGE ANY OWNERSHIP RIGHTS OR ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR
 * FITNESS FOR A PARTICULAR PURPOSE.
 * 
 * Copyright policy reference (version 1, approved 2009, updated 2011):
 * https://www.ecma-international.org/wp-content/uploads/Ecma_text_copyright_policy_version_1.pdf
 */
export const presetDefinitions: Readonly<Record<string, PresetDefinition>> = {
  "rect": {
    "adjustments": [],
    "guides": [],
    "text": [
      "l",
      "t",
      "r",
      "b"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "0",
        "r",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "r",
            "t"
          ],
          [
            "L",
            "r",
            "b"
          ],
          [
            "L",
            "l",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "roundRect": {
    "adjustments": [
      [
        "adj",
        "val 16667"
      ]
    ],
    "guides": [
      [
        "a",
        "pin 0 adj 50000"
      ],
      [
        "x1",
        "*/ ss a 100000"
      ],
      [
        "x2",
        "+- r 0 x1"
      ],
      [
        "y2",
        "+- b 0 x1"
      ],
      [
        "il",
        "*/ x1 29289 100000"
      ],
      [
        "ir",
        "+- r 0 il"
      ],
      [
        "ib",
        "+- b 0 il"
      ]
    ],
    "text": [
      "il",
      "il",
      "ir",
      "ib"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "0",
        "r",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "x1"
          ],
          [
            "A",
            "x1",
            "x1",
            "cd2",
            "cd4"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "A",
            "x1",
            "x1",
            "3cd4",
            "cd4"
          ],
          [
            "L",
            "r",
            "y2"
          ],
          [
            "A",
            "x1",
            "x1",
            "0",
            "cd4"
          ],
          [
            "L",
            "x1",
            "b"
          ],
          [
            "A",
            "x1",
            "x1",
            "cd4",
            "cd4"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "ellipse": {
    "adjustments": [],
    "guides": [
      [
        "idx",
        "cos wd2 2700000"
      ],
      [
        "idy",
        "sin hd2 2700000"
      ],
      [
        "il",
        "+- hc 0 idx"
      ],
      [
        "ir",
        "+- hc idx 0"
      ],
      [
        "it",
        "+- vc 0 idy"
      ],
      [
        "ib",
        "+- vc idy 0"
      ]
    ],
    "text": [
      "il",
      "it",
      "ir",
      "ib"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "3cd4",
        "il",
        "it"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "cd4",
        "il",
        "ib"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "cd4",
        "ir",
        "ib"
      ],
      [
        "0",
        "r",
        "vc"
      ],
      [
        "3cd4",
        "ir",
        "it"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "vc"
          ],
          [
            "A",
            "wd2",
            "hd2",
            "cd2",
            "cd4"
          ],
          [
            "A",
            "wd2",
            "hd2",
            "3cd4",
            "cd4"
          ],
          [
            "A",
            "wd2",
            "hd2",
            "0",
            "cd4"
          ],
          [
            "A",
            "wd2",
            "hd2",
            "cd4",
            "cd4"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "triangle": {
    "adjustments": [
      [
        "adj",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "a",
        "pin 0 adj 100000"
      ],
      [
        "x1",
        "*/ w a 200000"
      ],
      [
        "x2",
        "*/ w a 100000"
      ],
      [
        "x3",
        "+- x1 wd2 0"
      ]
    ],
    "text": [
      "x1",
      "vc",
      "x3",
      "b"
    ],
    "connections": [
      [
        "3cd4",
        "x2",
        "t"
      ],
      [
        "cd2",
        "x1",
        "vc"
      ],
      [
        "cd4",
        "l",
        "b"
      ],
      [
        "cd4",
        "x2",
        "b"
      ],
      [
        "cd4",
        "r",
        "b"
      ],
      [
        "0",
        "x3",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "b"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "r",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "diamond": {
    "adjustments": [],
    "guides": [
      [
        "ir",
        "*/ w 3 4"
      ],
      [
        "ib",
        "*/ h 3 4"
      ]
    ],
    "text": [
      "wd4",
      "hd4",
      "ir",
      "ib"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "0",
        "r",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "vc"
          ],
          [
            "L",
            "hc",
            "t"
          ],
          [
            "L",
            "r",
            "vc"
          ],
          [
            "L",
            "hc",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "parallelogram": {
    "adjustments": [
      [
        "adj",
        "val 25000"
      ]
    ],
    "guides": [
      [
        "maxAdj",
        "*/ 100000 w ss"
      ],
      [
        "a",
        "pin 0 adj maxAdj"
      ],
      [
        "x1",
        "*/ ss a 200000"
      ],
      [
        "x2",
        "*/ ss a 100000"
      ],
      [
        "x6",
        "+- r 0 x1"
      ],
      [
        "x5",
        "+- r 0 x2"
      ],
      [
        "x3",
        "*/ x5 1 2"
      ],
      [
        "x4",
        "+- r 0 x3"
      ],
      [
        "il",
        "*/ wd2 a maxAdj"
      ],
      [
        "q1",
        "*/ 5 a maxAdj"
      ],
      [
        "q2",
        "+/ 1 q1 12"
      ],
      [
        "il",
        "*/ q2 w 1"
      ],
      [
        "it",
        "*/ q2 h 1"
      ],
      [
        "ir",
        "+- r 0 il"
      ],
      [
        "ib",
        "+- b 0 it"
      ],
      [
        "q3",
        "*/ h hc x2"
      ],
      [
        "y1",
        "pin 0 q3 h"
      ],
      [
        "y2",
        "+- b 0 y1"
      ]
    ],
    "text": [
      "il",
      "it",
      "ir",
      "ib"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "y2"
      ],
      [
        "3cd4",
        "x4",
        "t"
      ],
      [
        "0",
        "x6",
        "vc"
      ],
      [
        "cd4",
        "x3",
        "b"
      ],
      [
        "cd4",
        "hc",
        "y1"
      ],
      [
        "cd2",
        "x1",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "b"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "r",
            "t"
          ],
          [
            "L",
            "x5",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "trapezoid": {
    "adjustments": [
      [
        "adj",
        "val 25000"
      ]
    ],
    "guides": [
      [
        "maxAdj",
        "*/ 50000 w ss"
      ],
      [
        "a",
        "pin 0 adj maxAdj"
      ],
      [
        "x1",
        "*/ ss a 200000"
      ],
      [
        "x2",
        "*/ ss a 100000"
      ],
      [
        "x3",
        "+- r 0 x2"
      ],
      [
        "x4",
        "+- r 0 x1"
      ],
      [
        "il",
        "*/ wd3 a maxAdj"
      ],
      [
        "it",
        "*/ hd3 a maxAdj"
      ],
      [
        "ir",
        "+- r 0 il"
      ]
    ],
    "text": [
      "il",
      "it",
      "ir",
      "b"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "cd2",
        "x1",
        "vc"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "0",
        "x4",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "b"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "x3",
            "t"
          ],
          [
            "L",
            "r",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "pentagon": {
    "adjustments": [
      [
        "hf",
        "val 105146"
      ],
      [
        "vf",
        "val 110557"
      ]
    ],
    "guides": [
      [
        "swd2",
        "*/ wd2 hf 100000"
      ],
      [
        "shd2",
        "*/ hd2 vf 100000"
      ],
      [
        "svc",
        "*/ vc  vf 100000"
      ],
      [
        "dx1",
        "cos swd2 1080000"
      ],
      [
        "dx2",
        "cos swd2 18360000"
      ],
      [
        "dy1",
        "sin shd2 1080000"
      ],
      [
        "dy2",
        "sin shd2 18360000"
      ],
      [
        "x1",
        "+- hc 0 dx1"
      ],
      [
        "x2",
        "+- hc 0 dx2"
      ],
      [
        "x3",
        "+- hc dx2 0"
      ],
      [
        "x4",
        "+- hc dx1 0"
      ],
      [
        "y1",
        "+- svc 0 dy1"
      ],
      [
        "y2",
        "+- svc 0 dy2"
      ],
      [
        "it",
        "*/ y1 dx2 dx1"
      ]
    ],
    "text": [
      "x2",
      "it",
      "x3",
      "y2"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "cd2",
        "x1",
        "y1"
      ],
      [
        "cd4",
        "x2",
        "y2"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "cd4",
        "x3",
        "y2"
      ],
      [
        "0",
        "x4",
        "y1"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "x1",
            "y1"
          ],
          [
            "L",
            "hc",
            "t"
          ],
          [
            "L",
            "x4",
            "y1"
          ],
          [
            "L",
            "x3",
            "y2"
          ],
          [
            "L",
            "x2",
            "y2"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "hexagon": {
    "adjustments": [
      [
        "adj",
        "val 25000"
      ],
      [
        "vf",
        "val 115470"
      ]
    ],
    "guides": [
      [
        "maxAdj",
        "*/ 50000 w ss"
      ],
      [
        "a",
        "pin 0 adj maxAdj"
      ],
      [
        "shd2",
        "*/ hd2 vf 100000"
      ],
      [
        "x1",
        "*/ ss a 100000"
      ],
      [
        "x2",
        "+- r 0 x1"
      ],
      [
        "dy1",
        "sin shd2 3600000"
      ],
      [
        "y1",
        "+- vc 0 dy1"
      ],
      [
        "y2",
        "+- vc dy1 0"
      ],
      [
        "q1",
        "*/ maxAdj -1 2"
      ],
      [
        "q2",
        "+- a q1 0"
      ],
      [
        "q3",
        "?: q2 4 2"
      ],
      [
        "q4",
        "?: q2 3 2"
      ],
      [
        "q5",
        "?: q2 q1 0"
      ],
      [
        "q6",
        "+/ a q5 q1"
      ],
      [
        "q7",
        "*/ q6 q4 -1"
      ],
      [
        "q8",
        "+- q3 q7 0"
      ],
      [
        "il",
        "*/ w q8 24"
      ],
      [
        "it",
        "*/ h q8 24"
      ],
      [
        "ir",
        "+- r 0 il"
      ],
      [
        "ib",
        "+- b 0 it"
      ]
    ],
    "text": [
      "il",
      "it",
      "ir",
      "ib"
    ],
    "connections": [
      [
        "0",
        "r",
        "vc"
      ],
      [
        "cd4",
        "x2",
        "y2"
      ],
      [
        "cd4",
        "x1",
        "y2"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "3cd4",
        "x1",
        "y1"
      ],
      [
        "3cd4",
        "x2",
        "y1"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "vc"
          ],
          [
            "L",
            "x1",
            "y1"
          ],
          [
            "L",
            "x2",
            "y1"
          ],
          [
            "L",
            "r",
            "vc"
          ],
          [
            "L",
            "x2",
            "y2"
          ],
          [
            "L",
            "x1",
            "y2"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "rightArrow": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ],
      [
        "adj2",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "maxAdj2",
        "*/ 100000 w ss"
      ],
      [
        "a1",
        "pin 0 adj1 100000"
      ],
      [
        "a2",
        "pin 0 adj2 maxAdj2"
      ],
      [
        "dx1",
        "*/ ss a2 100000"
      ],
      [
        "x1",
        "+- r 0 dx1"
      ],
      [
        "dy1",
        "*/ h a1 200000"
      ],
      [
        "y1",
        "+- vc 0 dy1"
      ],
      [
        "y2",
        "+- vc dy1 0"
      ],
      [
        "dx2",
        "*/ y1 dx1 hd2"
      ],
      [
        "x2",
        "+- x1 dx2 0"
      ]
    ],
    "text": [
      "l",
      "y1",
      "x2",
      "y2"
    ],
    "connections": [
      [
        "3cd4",
        "x1",
        "t"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "cd4",
        "x1",
        "b"
      ],
      [
        "0",
        "r",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "y1"
          ],
          [
            "L",
            "x1",
            "y1"
          ],
          [
            "L",
            "x1",
            "t"
          ],
          [
            "L",
            "r",
            "vc"
          ],
          [
            "L",
            "x1",
            "b"
          ],
          [
            "L",
            "x1",
            "y2"
          ],
          [
            "L",
            "l",
            "y2"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "leftArrow": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ],
      [
        "adj2",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "maxAdj2",
        "*/ 100000 w ss"
      ],
      [
        "a1",
        "pin 0 adj1 100000"
      ],
      [
        "a2",
        "pin 0 adj2 maxAdj2"
      ],
      [
        "dx2",
        "*/ ss a2 100000"
      ],
      [
        "x2",
        "+- l dx2 0"
      ],
      [
        "dy1",
        "*/ h a1 200000"
      ],
      [
        "y1",
        "+- vc 0 dy1"
      ],
      [
        "y2",
        "+- vc dy1 0"
      ],
      [
        "dx1",
        "*/ y1 dx2 hd2"
      ],
      [
        "x1",
        "+- x2  0 dx1"
      ]
    ],
    "text": [
      "x1",
      "y1",
      "r",
      "y2"
    ],
    "connections": [
      [
        "3cd4",
        "x2",
        "t"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "cd4",
        "x2",
        "b"
      ],
      [
        "0",
        "r",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "vc"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "x2",
            "y1"
          ],
          [
            "L",
            "r",
            "y1"
          ],
          [
            "L",
            "r",
            "y2"
          ],
          [
            "L",
            "x2",
            "y2"
          ],
          [
            "L",
            "x2",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "downArrow": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ],
      [
        "adj2",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "maxAdj2",
        "*/ 100000 h ss"
      ],
      [
        "a1",
        "pin 0 adj1 100000"
      ],
      [
        "a2",
        "pin 0 adj2 maxAdj2"
      ],
      [
        "dy1",
        "*/ ss a2 100000"
      ],
      [
        "y1",
        "+- b 0 dy1"
      ],
      [
        "dx1",
        "*/ w a1 200000"
      ],
      [
        "x1",
        "+- hc 0 dx1"
      ],
      [
        "x2",
        "+- hc dx1 0"
      ],
      [
        "dy2",
        "*/ x1 dy1 wd2"
      ],
      [
        "y2",
        "+- y1 dy2 0"
      ]
    ],
    "text": [
      "x1",
      "t",
      "x2",
      "y2"
    ],
    "connections": [
      [
        "3cd4",
        "hc",
        "t"
      ],
      [
        "cd2",
        "l",
        "y1"
      ],
      [
        "cd4",
        "hc",
        "b"
      ],
      [
        "0",
        "r",
        "y1"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "y1"
          ],
          [
            "L",
            "x1",
            "y1"
          ],
          [
            "L",
            "x1",
            "t"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "x2",
            "y1"
          ],
          [
            "L",
            "r",
            "y1"
          ],
          [
            "L",
            "hc",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "leftRightArrow": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ],
      [
        "adj2",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "maxAdj2",
        "*/ 50000 w ss"
      ],
      [
        "a1",
        "pin 0 adj1 100000"
      ],
      [
        "a2",
        "pin 0 adj2 maxAdj2"
      ],
      [
        "x2",
        "*/ ss a2 100000"
      ],
      [
        "x3",
        "+- r 0 x2"
      ],
      [
        "dy",
        "*/ h a1 200000"
      ],
      [
        "y1",
        "+- vc 0 dy"
      ],
      [
        "y2",
        "+- vc dy 0"
      ],
      [
        "dx1",
        "*/ y1 x2 hd2"
      ],
      [
        "x1",
        "+- x2 0 dx1"
      ],
      [
        "x4",
        "+- x3 dx1 0"
      ]
    ],
    "text": [
      "x1",
      "y1",
      "x4",
      "y2"
    ],
    "connections": [
      [
        "0",
        "r",
        "vc"
      ],
      [
        "cd4",
        "x3",
        "b"
      ],
      [
        "cd4",
        "x2",
        "b"
      ],
      [
        "cd2",
        "l",
        "vc"
      ],
      [
        "3cd4",
        "x2",
        "t"
      ],
      [
        "3cd4",
        "x3",
        "t"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "vc"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "x2",
            "y1"
          ],
          [
            "L",
            "x3",
            "y1"
          ],
          [
            "L",
            "x3",
            "t"
          ],
          [
            "L",
            "r",
            "vc"
          ],
          [
            "L",
            "x3",
            "b"
          ],
          [
            "L",
            "x3",
            "y2"
          ],
          [
            "L",
            "x2",
            "y2"
          ],
          [
            "L",
            "x2",
            "b"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "chevron": {
    "adjustments": [
      [
        "adj",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "maxAdj",
        "*/ 100000 w ss"
      ],
      [
        "a",
        "pin 0 adj maxAdj"
      ],
      [
        "x1",
        "*/ ss a 100000"
      ],
      [
        "x2",
        "+- r 0 x1"
      ],
      [
        "x3",
        "*/ x2 1 2"
      ],
      [
        "dx",
        "+- x2 0 x1"
      ],
      [
        "il",
        "?: dx x1 l"
      ],
      [
        "ir",
        "?: dx x2 r"
      ]
    ],
    "text": [
      "il",
      "t",
      "ir",
      "b"
    ],
    "connections": [
      [
        "3cd4",
        "x3",
        "t"
      ],
      [
        "cd2",
        "x1",
        "vc"
      ],
      [
        "cd4",
        "x3",
        "b"
      ],
      [
        "0",
        "r",
        "vc"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "x2",
            "t"
          ],
          [
            "L",
            "r",
            "vc"
          ],
          [
            "L",
            "x2",
            "b"
          ],
          [
            "L",
            "l",
            "b"
          ],
          [
            "L",
            "x1",
            "vc"
          ],
          [
            "Z"
          ]
        ]
      }
    ]
  },
  "line": {
    "adjustments": [],
    "guides": [],
    "text": null,
    "connections": [
      [
        "cd4",
        "l",
        "t"
      ],
      [
        "3cd4",
        "r",
        "b"
      ]
    ],
    "paths": [
      {
        "fill": true,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "r",
            "b"
          ]
        ]
      }
    ]
  },
  "straightConnector1": {
    "adjustments": [],
    "guides": [],
    "text": [
      "l",
      "t",
      "r",
      "b"
    ],
    "connections": [],
    "paths": [
      {
        "fill": false,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "r",
            "b"
          ]
        ]
      }
    ]
  },
  "bentConnector2": {
    "adjustments": [],
    "guides": [],
    "text": [
      "l",
      "t",
      "r",
      "b"
    ],
    "connections": [],
    "paths": [
      {
        "fill": false,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "r",
            "t"
          ],
          [
            "L",
            "r",
            "b"
          ]
        ]
      }
    ]
  },
  "bentConnector3": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "x1",
        "*/ w adj1 100000"
      ]
    ],
    "text": [
      "l",
      "t",
      "r",
      "b"
    ],
    "connections": [],
    "paths": [
      {
        "fill": false,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "x1",
            "t"
          ],
          [
            "L",
            "x1",
            "b"
          ],
          [
            "L",
            "r",
            "b"
          ]
        ]
      }
    ]
  },
  "bentConnector4": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ],
      [
        "adj2",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "x1",
        "*/ w adj1 100000"
      ],
      [
        "x2",
        "+/ x1 r 2"
      ],
      [
        "y2",
        "*/ h adj2 100000"
      ],
      [
        "y1",
        "+/ t y2 2"
      ]
    ],
    "text": [
      "l",
      "t",
      "r",
      "b"
    ],
    "connections": [],
    "paths": [
      {
        "fill": false,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "x1",
            "t"
          ],
          [
            "L",
            "x1",
            "y2"
          ],
          [
            "L",
            "r",
            "y2"
          ],
          [
            "L",
            "r",
            "b"
          ]
        ]
      }
    ]
  },
  "bentConnector5": {
    "adjustments": [
      [
        "adj1",
        "val 50000"
      ],
      [
        "adj2",
        "val 50000"
      ],
      [
        "adj3",
        "val 50000"
      ]
    ],
    "guides": [
      [
        "x1",
        "*/ w adj1 100000"
      ],
      [
        "x3",
        "*/ w adj3 100000"
      ],
      [
        "x2",
        "+/ x1 x3 2"
      ],
      [
        "y2",
        "*/ h adj2 100000"
      ],
      [
        "y1",
        "+/ t y2 2"
      ],
      [
        "y3",
        "+/ b y2 2"
      ]
    ],
    "text": [
      "l",
      "t",
      "r",
      "b"
    ],
    "connections": [],
    "paths": [
      {
        "fill": false,
        "stroke": true,
        "commands": [
          [
            "M",
            "l",
            "t"
          ],
          [
            "L",
            "x1",
            "t"
          ],
          [
            "L",
            "x1",
            "y2"
          ],
          [
            "L",
            "x3",
            "y2"
          ],
          [
            "L",
            "x3",
            "b"
          ],
          [
            "L",
            "r",
            "b"
          ]
        ]
      }
    ]
  }
};
