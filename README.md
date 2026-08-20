# safari-data-pipelines
Python data pipelines for scheduled data extraction, transformation, reporting, and automation workflows in Safari.

## Basic Pipeline Scaffold

A reusable scaffold is available under `pipelines/scaffold/` for quickly creating new ETL-style pipelines.

- `pipelines/scaffold/base_pipeline.py`
	- `BasePipeline` with `extract()`, `transform()`, `load()` lifecycle
	- Standardized logging and runtime summary
	- Config merge helper: defaults + `config/config.yaml` + `config/auth.yaml`
- `pipelines/scaffold/example_basic_pipeline.py`
	- Minimal working example using pandas
	- Good starting point to copy for real pipelines

Run the example:

```bash
python -m pipelines.scaffold.example_basic_pipeline
```

Optional config shape in YAML files:

```yaml
pipelines:
	basic_scaffold_pipeline:
		transform:
			uppercase_name: true
		load:
			target: console
```
