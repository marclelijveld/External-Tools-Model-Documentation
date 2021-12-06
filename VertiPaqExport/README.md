# VertiPaq-Analyzer

VertiPaq Analyzer 2.0 is an open source project available on GitHub that includes several libraries made available on NuGet, enabling the integration of VertiPaq Analyzer features in different tools. These are the libraries available:

- **Dax.Metadata** is a representation of the Tabular model including additional information from DMV and statistics about data distribution in the model. This is a .NET Standard library without dependencies on the Tabular Object Model (TOM). This is the core of VertiPaq Analyzer and includes all the information about the data model in a format that will allow a complete anonymization of the data model in a future release.
- **Dax.Model.Extractor** populates the Dax.Metadata.Model extracting data from a Tabular model. This library depends on TOM and .NET Framework 2.7.1.
- **Dax.ViewModel** provides a view over Dax.Metadata.Model data that can be integrated in applications using the view model, such as the one implemented in DAX Studio (use DAX Studio 2.9.3 or later version). This library is a .NET Standard library.
- **Dax.ViewVpaExport** exports the Dax.Model for the consumption in VertiPaq Analyzer 2.0 client tools. This library is a .NET Standard library.
- **Dax.Vpax** supports the VPAX format.

Using the VertiPaq Analyzer libraries, a client tool can easily extract the information from a Tabular model, export them in a standard file, and read the same information for visualization purposes.
VPAX Format

The VPAX file format is a ZIP file containing the following files:
- **DaxModel.json** is a serialization of the Dax.Metadata.Model.
- **DaxVpaView.json** is a serialization of the Dax.ViewVpaExport that is used to import data in the VertiPaq Analyzer 2.0 Excel file using a specific macro.
- **Model.bim** is an optional file that exports the complete model in TOM format. This file is not used by current client tools, but it could be useful for future extensions.
