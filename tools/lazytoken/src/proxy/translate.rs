use chrono::Utc;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};

#[derive(Debug, Deserialize)]
pub struct OpenAiChatCompletionRequest {
    pub model: String,
    pub messages: Vec<OpenAiMessage>,
    pub stream: Option<bool>,
}

#[derive(Debug, Deserialize)]
pub struct OpenAiMessage {
    pub role: String,
    pub content: Value,
}

#[derive(Debug, Serialize)]
pub struct PlatformChatRequest {
    pub session_id: u64,
    pub model: String,
    pub messages: Vec<PlatformMessage>,
}

#[derive(Debug, Serialize, PartialEq, Eq)]
pub struct PlatformMessage {
    pub id: u64,
    pub role: String,
    pub content: String,
}

pub fn to_platform_request(
    req: &OpenAiChatCompletionRequest,
    session_id: u64,
) -> Result<PlatformChatRequest, String> {
    let messages = req
        .messages
        .iter()
        .enumerate()
        .map(|(idx, msg)| {
            Ok(PlatformMessage {
                id: (idx as u64) + 1,
                role: msg.role.clone(),
                content: extract_text_content(&msg.content)?,
            })
        })
        .collect::<Result<Vec<_>, String>>()?;

    Ok(PlatformChatRequest {
        session_id,
        model: req.model.clone(),
        messages,
    })
}

pub fn extract_text_content(content: &Value) -> Result<String, String> {
    match content {
        Value::String(text) => Ok(text.clone()),
        Value::Array(parts) => {
            let mut chunks = Vec::new();
            for part in parts {
                let part_type = part.get("type").and_then(Value::as_str).unwrap_or_default();
                if part_type != "text" {
                    return Err("当前 MVP 仅支持文本消息 content（不支持图片/视频）".to_string());
                }
                let text = part
                    .get("text")
                    .and_then(Value::as_str)
                    .ok_or_else(|| "text content 缺失 text 字段".to_string())?;
                chunks.push(text.to_string());
            }
            Ok(chunks.join(""))
        }
        _ => Err("messages[*].content 必须是字符串或 text 数组".to_string()),
    }
}

pub fn build_non_stream_response(model: &str, completion_text: &str) -> Value {
    let now = Utc::now();
    json!({
        "id": format!("chatcmpl-{}", now.timestamp_millis()),
        "object": "chat.completion",
        "created": now.timestamp(),
        "model": model,
        "choices": [
            {
                "index": 0,
                "message": {
                    "role": "assistant",
                    "content": completion_text,
                },
                "finish_reason": "stop"
            }
        ],
        "usage": {
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "total_tokens": 0
        }
    })
}

pub fn build_stream_chunk(model: &str, chunk_id: &str, content: &str) -> Value {
    json!({
        "id": chunk_id,
        "object": "chat.completion.chunk",
        "created": Utc::now().timestamp(),
        "model": model,
        "choices": [
            {
                "index": 0,
                "delta": {
                    "content": content
                },
                "finish_reason": Value::Null
            }
        ]
    })
}

pub fn build_stream_stop_chunk(model: &str, chunk_id: &str) -> Value {
    json!({
        "id": chunk_id,
        "object": "chat.completion.chunk",
        "created": Utc::now().timestamp(),
        "model": model,
        "choices": [
            {
                "index": 0,
                "delta": {},
                "finish_reason": "stop"
            }
        ]
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn to_platform_request_assigns_incremental_message_ids() {
        let req = OpenAiChatCompletionRequest {
            model: "gpt-5.4".to_string(),
            messages: vec![
                OpenAiMessage {
                    role: "system".to_string(),
                    content: json!("sys"),
                },
                OpenAiMessage {
                    role: "user".to_string(),
                    content: json!("hello"),
                },
            ],
            stream: Some(false),
        };

        let payload = to_platform_request(&req, 42).expect("platform request");
        assert_eq!(payload.session_id, 42);
        assert_eq!(payload.messages.len(), 2);
        assert_eq!(payload.messages[0].id, 1);
        assert_eq!(payload.messages[1].id, 2);
        assert_eq!(payload.messages[1].content, "hello");
    }

    #[test]
    fn extract_text_content_accepts_openai_text_parts() {
        let content = json!([
            {"type":"text","text":"你"},
            {"type":"text","text":"好"}
        ]);

        let text = extract_text_content(&content).expect("text parts should be supported");
        assert_eq!(text, "你好");
    }

    #[test]
    fn extract_text_content_rejects_non_text_parts() {
        let content = json!([{"type":"input_image","image_url":"x"}]);
        let err = extract_text_content(&content).expect_err("image should be rejected in MVP");
        assert!(err.contains("不支持图片/视频"));
    }
}
