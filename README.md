# Calendar Schedule PCF Control

A dataset-based calendar control that can render appointments in a month/week/day layout inside Power Platform Canvas and Model-driven apps.

## Overview

## Table of contents
1. [Prerequisites](#prerequisites)
2. [Setup](#setup)
3. [Control manifest wiring](#control-manifest-wiring)
4. [Field mapping reference](#field-mapping-reference)
5. [Using the control inside apps](#using-the-control-inside-apps)
6. [Build & deployment](#build--deployment)
7. [Testing notes](#testing-notes)

## Included solution

The solution `calenderPCF` is included in the project root (`calenderPCF/`) and can be imported into your Power Platform environment to install and use the control directly.
# Calendar Schedule — Power Platform PCF Control

This repository contains a dataset-based Power Apps Component Framework (PCF) control that renders Dataverse appointment records in month/week/day views for Power Apps (Canvas and model-driven experiences).

## Overview

- **Control type:** Dataset PCF control (consumes a Dataverse view / dataset exposed by the host).
- **Status:** Tested in Canvas apps (verified). Model-driven apps are supported via the same dataset contract — validate in your target environment.
- **Main source file:** `Calender.tsx` (normalizes dataset records, computes layout, renders UI).

## Contents

- Prerequisites
- Install & run
- Manifest & field mappings
- Configure in Canvas and model-driven apps
- Build, package, and deploy
- Testing & troubleshooting
- Resources

## Prerequisites

- Power Platform CLI (`pac`) and PCF toolchain available to package/push controls.
- Node.js and `npm` (install dependencies with `npm install`).
- A Dataverse table (entity) and a view that exposes the columns required by the control.

## Install & run

1. From the repository root install dependencies:

```bash
npm install
```

2. Common scripts (from `package.json`):

- `npm run build` — compile and bundle the control
- `npm run start` — run dev server with live reload for development
- `npm run lint` — run linting

## Manifest & field mappings

The control consumes a dataset provided by the host. Confirm the manifest (`ControlManifest.Input.xml`) includes a `<data-set>` declaration and exposes the input properties below.

### Required property mappings (set these to the logical column names in your Dataverse view):

- `fromColumn` — Start DateTime column (example: `scheduledstart`, `msdyn_start`).
- `toColumn` — End DateTime column (example: `scheduledend`, `msdyn_end`).
- `titleColumn` — Primary title (example: `subject` or `name`).
- `subtitleColumn` — Secondary text (owner name, location, etc.).
- `typeColumn` — Category/type column (option set label or text) used for styling.

### Implementation notes

- The control calls `record.getValue(<columnLogicalName>)` for each mapped property. If the mapped column isn't present in the view the record may be skipped or shown with fallback values.
- Ensure the Dataverse view that feeds the subgrid/dataset includes the mapped columns (add them in the view editor or via FetchXML).

## Configure in Canvas and model-driven apps

### Canvas apps

1. Add the PCF control to a dataset area or gallery bound to a Dataverse view.
2. Open the control properties and set the five column mappings to the view's logical names.

### Model-driven apps (subgrid / view)

1. Add the custom control to a subgrid or dataset-capable field on the form.
2. Configure the property mappings for the control so it reads the correct Dataverse columns.

## Build, package, and deploy

1. Build the control:

```bash
npm run build
```

2. Package or push the control to an environment. Example using `pac`:

```bash
# Push PCF control to your environment (example)
pac pcf push --solution-unique-name <ExampleSolution>
```



