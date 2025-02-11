---
title: "Function Registration"
---
**Note:** This document reflects changes made in Excel-DNA v1.9.

Excel-DNA support the creation of user-defined functions for Excel in .NET. This document describes how functions are selected for registration, supported method signatures, method conversions and extension points.




### Changes from earlier versions
In v1.9 we incorporate the functionality previously exposes in the separate `ExcelDna.Registration` library (and package) into the main Excel-DNA library.
To use the extended registration features under older versions, explict registration was required.
Under v1.9 we have:
* expanded the supported parameter and return types for functions markes with `[ExcelFunction]`.
* added support for async and streaming functions and object handles (with the `[return:ExcelHandle]` and `[ExcelHandle]` attributes) in the main library
* migrated registration extension points like `FunctionExecutionHandler` from the ExcelDna.Registration package to the main ExcelDna.Integration library.
