#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SseEvent {
    pub event: String,
    pub data: String,
}

pub fn parse_events(payload: &str) -> Vec<SseEvent> {
    let normalized = payload.replace("\r\n", "\n");
    normalized
        .split("\n\n")
        .filter_map(parse_event_block)
        .collect()
}

pub fn extract_content_chunks(payload: &str) -> Vec<String> {
    parse_events(payload)
        .into_iter()
        .filter(|e| e.event == "content")
        .map(|e| e.data)
        .collect()
}

fn parse_event_block(block: &str) -> Option<SseEvent> {
    let mut event_name = String::new();
    let mut data_lines = Vec::new();

    for line in block.lines() {
        if let Some((key, value)) = line.split_once(':') {
            let value = value.trim_start();
            match key {
                "event" => event_name = value.to_string(),
                "data" => data_lines.push(value.to_string()),
                _ => {}
            }
        }
    }

    if event_name.is_empty() {
        return None;
    }

    Some(SseEvent {
        event: event_name,
        data: data_lines.join("\n"),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_events_reads_content_and_done() {
        let raw =
            "event: content\ndata: 你\n\nevent: content\ndata: 好\n\nevent: done\ndata: done\n\n";
        let events = parse_events(raw);
        assert_eq!(events.len(), 3);
        assert_eq!(events[0].event, "content");
        assert_eq!(events[1].data, "好");
        assert_eq!(events[2].event, "done");
    }

    #[test]
    fn extract_content_chunks_ignores_non_content_events() {
        let raw = "event: ai_message_id\ndata: 999\n\nevent: content\ndata: hello\n\nevent: done\ndata: done\n\n";
        let chunks = extract_content_chunks(raw);
        assert_eq!(chunks, vec!["hello".to_string()]);
    }
}
