# Service templates and service costs

A service template is the standard job for one machine and one service interval, for example "A300AM 500 h service". It lists:

- **Materials:** the service kit, filters, parts and oils, with quantities.
- **Labour:** the standard hours and labour rate.

IronLog prices a template at current store costs. That price becomes the forecast cost of that machine's upcoming service in the weekly forum, the weekly plan, maintenance insights and the upcoming-cost reports.

## Building templates from history

Go to Maintenance → Service templates → **Build templates from service history**, and click **Build proposals**. For every machine and service interval in the maintenance plans, IronLog proposes a template:

1. **From service history.** It looks at past service work orders for that plan and what stores issued to them. Items issued on at least half of those services are kept, at their typical (median) quantity. Returns to store are subtracted, and one-off extras are left out.
2. **From a service kit.** Without history, it uses the kit whose stock code names the machine and interval (for example `TEICH-A300AM-500HRS`). Failing that, it uses a model kit whose name matches the machine's make and model and the interval (for example "CAT 350 500hr service kit"). It never borrows a kit coded to a different machine.
3. **Oils from a sibling.** A kit-only proposal copies the oil quantities from a machine of the same model whose template came from history.

**Labour** isn't captured today, so every template uses a standard time: 2 h up to 250 h, 4 h to 500 h, 6 h to 1,000 h, 8 h to 2,000 h, 12 h beyond, and 2 h for LDV services. The rate is the default labour rate in cost settings. Edit a template (new revision) to change either.

Tick the proposals you want and click **Create ticked templates**. Each one is assigned to its machine. If the machine already had its own template for that interval, the old one is retired but kept for history. Machines that already have their own template start unticked.

## Which cost wins

For each upcoming service, IronLog uses, in this order:

1. an all-in quote entered in the weekly forum inputs
2. manual parts and/or labour entered in the weekly forum inputs; whichever half is missing comes from the template
3. **the machine's service template**
4. the average of past services

Reports show the source in plain words: "service template", "past service average", "needs pricing" and so on.

The builder is `api/utils/serviceTemplateBuilder.js`, with tests in `api/test/serviceTemplateBuilder.test.js`.
