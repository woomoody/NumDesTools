use chrono::Local;
use serde::{Deserialize, Serialize};
use std::{fs, io, path::PathBuf};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Route {
    pub harness: String,
    pub scope: String,
    pub role: String,
    pub kind: String,
    pub model: String,
    pub effort: String,
}
#[derive(Debug, Serialize, Deserialize)]
pub struct RoutesFile {
    pub version: u32,
    pub routes: Vec<Route>,
}
#[derive(Debug, Serialize, Deserialize)]
pub struct Catalog {
    pub date: String,
    pub models: Vec<String>,
    #[serde(default)]
    pub info: std::collections::HashMap<String, ModelInfo>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ModelInfo {
    pub tier: String,
    pub csharp: String,
    pub input: String,
    pub output: String,
}

pub fn root() -> PathBuf {
    let exe_dir = std::env::current_exe()
        .ok()
        .and_then(|path| path.parent().map(PathBuf::from));
    if let Some(dir) = &exe_dir {
        if dir.join("lazymodel.json").exists() {
            return dir.clone();
        }
    }
    if let Ok(dir) = std::env::var("LAZYMODEL_HOME") {
        let path = PathBuf::from(dir);
        if path.join("lazymodel.json").exists() {
            return path;
        }
    }
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}
pub fn routes_path() -> PathBuf {
    root().join("lazymodel.json")
}
pub fn catalog_path() -> PathBuf {
    root().join("litellm_model_catalog.json")
}
pub fn load_routes() -> io::Result<RoutesFile> {
    serde_json::from_str(&fs::read_to_string(routes_path())?)
        .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))
}
pub fn save_routes(d: &RoutesFile) -> io::Result<()> {
    fs::write(
        routes_path(),
        serde_json::to_string_pretty(d).unwrap() + "\n",
    )
}
pub fn load_catalog() -> Catalog {
    fs::read_to_string(catalog_path())
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or(Catalog {
            date: String::new(),
            models: vec![],
            info: std::collections::HashMap::new(),
        })
}

pub fn model_info(_catalog: &Catalog, model: &str) -> ModelInfo {
    let home = PathBuf::from(std::env::var("USERPROFILE").unwrap_or_default());
    let score_text = fs::read_to_string(home.join(".claude/model_scores.md")).unwrap_or_default();
    let price_path = root().join("../../NumDesTools.Tests/model_prices.json");
    let prices: Vec<serde_json::Value> = fs::read_to_string(price_path)
        .ok()
        .and_then(|s| serde_json::from_str::<serde_json::Value>(&s).ok())
        .and_then(|v| v["prices"].as_array().cloned())
        .unwrap_or_default();
    let mut info = ModelInfo {
        tier: "-".into(),
        csharp: "-".into(),
        input: "-".into(),
        output: "-".into(),
    };
    if let Some(line) = score_text
        .lines()
        .find(|line| line.starts_with(&format!("| {} |", model)))
    {
        let cells: Vec<&str> = line.trim_matches('|').split('|').map(str::trim).collect();
        info.tier = cells.get(1).unwrap_or(&"-").to_string();
        info.csharp = cells.get(2).unwrap_or(&"-").to_string();
    }
    if let Some(price) = prices.iter().find(|price| {
        model.to_lowercase().starts_with(
            price["prefix"]
                .as_str()
                .unwrap_or("")
                .to_lowercase()
                .as_str(),
        )
    }) {
        info.input = price["input"].to_string();
        info.output = price["output"].to_string();
    }
    info
}

pub fn refresh_catalog() -> Catalog {
    let mut c = load_catalog();
    let today = Local::now().format("%Y-%m-%d").to_string();
    if c.date == today && !c.models.is_empty() {
        return c;
    }
    let Some(key) = std::env::var("LITELLM_API_KEY").ok().or_else(hermes_key) else {
        return c;
    };
    if let Ok(resp) = ureq::get("https://litellm.solotopia.net/v1/models")
        .set("Authorization", &format!("Bearer {key}"))
        .call()
    {
        if let Ok(v) = resp.into_json::<serde_json::Value>() {
            let mut m: Vec<String> = v["data"]
                .as_array()
                .into_iter()
                .flatten()
                .filter_map(|x| x["id"].as_str().map(str::to_string))
                .collect();
            m.sort_by_key(|x| x.to_lowercase());
            m.dedup();
            if !m.is_empty() {
                c = Catalog {
                    date: today,
                    models: m,
                    info: std::collections::HashMap::new(),
                };
                let _ = fs::write(
                    catalog_path(),
                    serde_json::to_string_pretty(&c).unwrap() + "\n",
                );
            }
        }
    }
    c
}
fn hermes_key() -> Option<String> {
    let p = PathBuf::from(std::env::var("LOCALAPPDATA").ok()?).join("hermes/config.yaml");
    fs::read_to_string(p).ok()?.lines().find_map(|l| {
        if l.trim_start().starts_with("api_key:") {
            Some(
                l.split_once(':')?
                    .1
                    .trim()
                    .trim_matches(['\'', '"'])
                    .to_string(),
            )
        } else {
            None
        }
    })
}

pub fn apply_routes(d: &RoutesFile) -> io::Result<PathBuf> {
    let home = PathBuf::from(std::env::var("USERPROFILE").unwrap_or_default());
    let local = PathBuf::from(std::env::var("LOCALAPPDATA").unwrap_or_default());
    let backup_dir = root().join(format!(
        "lazymodel-backup-{}",
        Local::now().format("%Y%m%d-%H%M%S")
    ));
    fs::create_dir_all(&backup_dir)?;
    for (scope, path) in [
        ("global", home.join(".config/opencode/opencode.jsonc")),
        ("CCDS", home.join("CCDS/opencode.jsonc")),
        ("CCglm", home.join("CCglm/opencode.jsonc")),
        ("CCKimi", home.join("CCKimi/opencode.jsonc")),
        ("CCGame", home.join("CCGame/.opencode/opencode.jsonc")),
    ] {
        if path.exists() {
            let mut v: serde_json::Value = serde_json::from_str(&fs::read_to_string(&path)?)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;
            if let Some(r) = find(d, "opencode", scope, "primary")
                .or_else(|| find(d, "opencode", "global", "primary"))
            {
                v["model"] = serde_json::Value::String(format!("litellm/{}", r.model));
                set_opencode_variant(&mut v, &r.model, &r.effort);
            }
            if let Some(r) = find(d, "opencode", "all", "small_model") {
                v["small_model"] = serde_json::Value::String(format!("litellm/{}", r.model));
            }
            backup(&path, &backup_dir)?;
            fs::write(path, serde_json::to_string_pretty(&v).unwrap() + "\n")?;
        }
    }
    for (scope, path) in [
        ("global", home.join(".config/opencode/oh-my-openagent.json")),
        ("CCGame", home.join("CCGame/.opencode/oh-my-openagent.json")),
    ] {
        if path.exists() {
            let mut v: serde_json::Value = serde_json::from_str(&fs::read_to_string(&path)?)
                .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;
            for sec in ["agents", "categories"] {
                if let Some(map) = v.get_mut(sec).and_then(|x| x.as_object_mut()) {
                    for (name, spec) in map {
                        let role = format!("{}.{}", sec, name);
                        let r = find(d, "oh-my-openagent", scope, &role)
                            .or_else(|| find(d, "oh-my-openagent", "global", &role));
                        if let Some(r) = r {
                            spec["model"] =
                                serde_json::Value::String(format!("litellm/{}", r.model));
                        }
                    }
                }
            }
            backup(&path, &backup_dir)?;
            fs::write(path, serde_json::to_string_pretty(&v).unwrap() + "\n")?;
        }
    }
    let agents = home.join(".claude/agents");
    if agents.exists() {
        for e in fs::read_dir(agents)? {
            let p = e?.path();
            if p.extension().and_then(|x| x.to_str()) != Some("md") {
                continue;
            }
            let role = p.file_stem().and_then(|x| x.to_str()).unwrap_or_default();
            if let Some(r) = find(d, "claude-code", "global", role) {
                let s = fs::read_to_string(&p)?;
                let s = s
                    .lines()
                    .map(|l| {
                        if l.starts_with("model:") {
                            format!("model: {}", r.model)
                        } else {
                            l.to_string()
                        }
                    })
                    .collect::<Vec<_>>()
                    .join("\n")
                    + "\n";
                backup(&p, &backup_dir)?;
                fs::write(p, s)?;
            }
        }
    }
    if let Some(r) = find(d, "hermes", "global", "model.default") {
        let p = local.join("hermes/config.yaml");
        if p.exists() {
            let s = fs::read_to_string(&p)?;
            let s = s
                .lines()
                .map(|l| {
                    if l.trim_start().starts_with("default:") {
                        format!("  default: {}", r.model)
                    } else {
                        l.to_string()
                    }
                })
                .collect::<Vec<_>>()
                .join("\n")
                + "\n";
            let s = set_hermes_effort(&s, &r.effort);
            backup(&p, &backup_dir)?;
            fs::write(p, s)?;
        }
    }
    Ok(backup_dir)
}
fn set_opencode_variant(v: &mut serde_json::Value, model: &str, effort: &str) {
    if effort == "default" {
        return;
    }
    if let Some(spec) = v["provider"]["litellm"]["models"].get_mut(model) {
        if !spec.get("variants").map_or(false, |x| x.is_object()) {
            spec["variants"] = serde_json::json!({});
        }
        spec["variants"][effort] = serde_json::json!({
            "reasoningEffort": effort,
            "textVerbosity": "low"
        });
    }
}

fn set_hermes_effort(text: &str, effort: &str) -> String {
    if effort == "default" {
        return text.to_string();
    }
    let mut lines: Vec<String> = text.lines().map(str::to_string).collect();
    if let Some(line) = lines
        .iter_mut()
        .find(|line| line.trim_start().starts_with("reasoning_effort:"))
    {
        *line = format!("  reasoning_effort: {}", effort);
    } else if let Some(i) = lines.iter().position(|line| line.trim() == "agent:") {
        lines.insert(i + 1, format!("  reasoning_effort: {}", effort));
    }
    lines.join("\n") + "\n"
}

fn find<'a>(d: &'a RoutesFile, h: &str, s: &str, r: &str) -> Option<&'a Route> {
    d.routes
        .iter()
        .find(|x| x.harness == h && x.scope == s && x.role == r)
}
fn backup(p: &PathBuf, b: &PathBuf) -> io::Result<()> {
    fs::copy(
        p,
        b.join(p.to_string_lossy().replace([':', '\\', '/'], "_")),
    )
    .map(|_| ())
}
