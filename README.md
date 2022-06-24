# Can I create custom Excel Linked Data Types with SPARQL?

## What are Excel Linked Data Types?

![linked-data-types](https://user-images.githubusercontent.com/8033981/174072492-b43a2d34-9c1a-497a-b5a5-3ad2beddaa8c.gif)

That's cool isn't it?

## Can I create my own data types using SPARQL?

Yes but todo exactly that you need a license.

```
Licensing
The Excel Data Types Gallery and connected experiences to Power BI featured tables is available for Excel subscribers with a Power BI Pro service plan.
```

Check

- https://docs.microsoft.com/en-us/power-bi/collaborate-share/service-excel-featured-tables
- https://docs.microsoft.com/en-us/power-bi/collaborate-share/service-create-excel-featured-tables

But hold on. We can still do Data Types without this license but they are not in the Company Data Catalog and are a bit less nice to use. I'll show you how todo it.

## Power Query

Power Query is a tool get Data From differnet sources and Transform it. Good news for Mac Users. Power Query is available for Mac. Bad news it does currently not support what we need in this example. Hope fully it will change it the future.

So you have to use Windows right now. But once you finished your Data Type it will work on a Mac as well.

### How it's done

https://user-images.githubusercontent.com/8033981/175544919-eab4482c-c87c-4893-91d2-d5845918ea81.mp4

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