# Shopping candidates from a meal plan

`build_shopping_candidates` turns the recipes planned in a date range into a
list of items to check before shopping. The household ticks off what it
already has, and Home Assistant moves the rest to the shopping list.

## Recipe note schema

Each recipe task's note is JSON. The tool reads these fields:

```json
{
  "dates_planned": ["2026-10-17"],
  "kategori": "hovedret",
  "kød": [{ "vare": ["Bacon i tern", "Bacon i skiver"], "mængde": "1 pakke" }],
  "ingredienser": [{ "vare": ["Hakkede tomater", "Flåede tomater"], "mængde": "2 dåser" }]
}
```

`vare` lists alternatives, with the preferred one first. `kød` holds the main
protein and is checked against the freezer list. `ingredienser` holds
everything else and is checked against the pantry list. A recipe with
`kategori: "rester"` needs nothing. Any other planned recipe that has no
`ingredienser` is listed under `recipes_missing_ingredients`.

## What happens to each item

| Situation | Result |
|---|---|
| Meat found in an open freezer task (`<Placering> – <Vare> – …`) | `from_freezer`, not bought |
| Ingredient whose name exactly matches an open pantry task | `skipped_pantry` |
| Already open on the shopping list | `already_on_shopping` |
| Already open on the staging list (rerun) | `already_staged` |
| Otherwise | A staging task is written, one per item, with its amounts and recipes combined |

Shopping lists keep their completed items, and their titles are free text
("Flåede tomater – 2 dåser"). Matching therefore compares a reduced key: the
part before the first dash, comma or parenthesis, without amounts or units,
and with simple Danish plural endings removed. A title may add one word to the
item ("Brune champignon" matches "Champignon"), but a more general title never
stands in for a more specific item ("Tomater" does not match "Hakkede
tomater"). Pantry matching is exact, so a pantry "Paprika" does not swallow
"Rød paprika".

When a completed shopping task matches, the staging note ends with
`shopping_id: <id>`. Home Assistant then reopens that task instead of adding a
duplicate.

## Home Assistant side

The `packages/indkob.yaml` package in the Home Assistant config does three
things:

- It shows a browser_mod popup on the kiosk when the staging list gets items.
- It adds a "Send resten til Shopping" script. For each open staging item, the
  script reopens `shopping_id` with `ms365_todo.complete_todo completed: false`,
  or adds a new item if there is no id. Then it removes the staging item.
- It closes the popup when the staging list is empty.

Home Assistant can't see completed shopping items, because `show_completed` is
off, so the id has to come from this tool. Graph task ids are the same in both
places.
