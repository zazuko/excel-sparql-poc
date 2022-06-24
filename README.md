# Custom Excel Linked Data Types using RDF & SPARQL

## What is Excel Linked Data Types?

Excel [linked data types](https://support.microsoft.com/en-us/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) connects an Excel sheet to external data sources. Even though it's called linked data types, it has nothing to do with what the RDF & Semantic Web community calls [linked data](https://en.wikipedia.org/wiki/Linked_data).

But we can connect external data sources directly to Excel. That sounds like something we should do with real linked data served from a public or company internal SPARQL endpoint! Check out this screencast:

![linked-data-types](https://user-images.githubusercontent.com/8033981/174072492-b43a2d34-9c1a-497a-b5a5-3ad2beddaa8c.gif)

This example is based on a data source provided by Microsoft. Once we declare the type, Excel offers to add additional columns that are fetched from an external data source and are linked to the type. It is not clear to the user where this data is coming from.

## So can we create our data types using SPARQL?

Yes! We could fetch open data available in RDF from public SPARQL endpoints or we can do that within our organization using an internal SPARQL endpoint. Unfortunately, this requires a license of Microsoft Power BI:


>**Licensing**
>
>The Excel Data Types Gallery and connected experiences to Power BI featured tables is available for Excel subscribers with a Power BI Pro service plan.


This is documented here:

- https://docs.microsoft.com/en-us/power-bi/collaborate-share/service-excel-featured-tables
- https://docs.microsoft.com/en-us/power-bi/collaborate-share/service-create-excel-featured-tables

If you would like to integrate this in your organization, we at [Zazuko](https://zazuko.com/) can support you. Please get in [contact](mailto:info@zazuko.com?subject=Excel Linked Data Types) with us.

## Power Query

Without Power BI we can still do data types, even though the integration is (a lot) less fancy. This is called [Power Query](https://support.microsoft.com/en-us/office/about-power-query-in-excel-7104fbee-9e62-4cb9-a02e-5bfb1a6c536a) and in the example sheet provided in this repository, there is an embedded SPARQL query.

Power Query is a tool to fetch data from different sources. The example can only be edited by the Windows version of Excel. The *latest* version of Office for Mac supports Power Query, unfortunately only for running it. Click the *Refresh All* button in the *Data* tab of Excel. If you get `UNKNOWN` in one of your cells, your version of Excel is not recent enough.

### How it's done

https://user-images.githubusercontent.com/8033981/175544919-eab4482c-c87c-4893-91d2-d5845918ea81.mp4

This is the SPARQL query used in the template for this example:

### Template for a POST request in Power Query

```vb
let
    URL = "https://lindas.admin.ch/query",
    SPARQL = "
PREFIX schema: <http://schema.org/>
PREFIX rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#>
PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>

SELECT * WHERE {
	<https://ld.admin.ch/department> schema:hasDefinedTerm ?term .
  	?term schema:alternateName ?altNameDE;
  		  schema:alternateName ?altNameIT;
  		  schema:alternateName ?altNameFR;
  		  schema:alternateName ?altNameRM;
      	  	  schema:name ?nameDE;
  		  schema:name ?nameIT;
  		  schema:name ?nameFR;
  		  schema:name ?nameRM.

    FILTER(langMatches(lang(?altNameDE), 'de'))
    FILTER(langMatches(lang(?altNameIT), 'it'))
    FILTER(langMatches(lang(?altNameFR), 'fr'))
    FILTER(langMatches(lang(?altNameRM), 'rm'))

    FILTER(langMatches(lang(?nameDE), 'de'))
    FILTER(langMatches(lang(?nameIT), 'it'))
    FILTER(langMatches(lang(?nameFR), 'fr'))
    FILTER(langMatches(lang(?nameRM), 'rm'))
}
    ",
    HEADERS = [#"Content-Type" = "application/x-www-form-urlencoded", #"Accept" = "text/csv"],

    Source = Csv.Document(
        Web.Contents(URL, [
            Headers = HEADERS,
            Content = Text.ToBinary("query="&Uri.EscapeDataString(SPARQL))
        ])
    )
   in
    Source
```

This example fetches abbreviations and labels for Swiss Government departments in all four official Languages in Switzerland.
