import yaml

def init(path="field_config.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)
