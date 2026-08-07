use crate::proxy::translate::PlatformChatRequest;

pub fn post_chat(
    base_url: &str,
    token: &str,
    request_body: &PlatformChatRequest,
) -> Result<String, String> {
    let url = format!("{}/api/chat", base_url.trim_end_matches('/'));
    let body =
        serde_json::to_string(request_body).map_err(|e| format!("序列化上游请求失败: {}", e))?;

    let agent = ureq::AgentBuilder::new()
        .timeout(std::time::Duration::from_secs(300))
        .build();

    match agent
        .post(&url)
        .set("Authorization", &format!("Bearer {}", token))
        .set("Accept", "text/event-stream")
        .set("Content-Type", "application/json")
        .send_string(&body)
    {
        Ok(resp) => {
            let status = resp.status();
            let text = resp.into_string().unwrap_or_default();
            if (200..300).contains(&status) {
                Ok(text)
            } else {
                Err(format!("上游 HTTP {}: {}", status, text))
            }
        }
        Err(e) => Err(format!("上游请求失败: {}", e)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::proxy::translate::PlatformMessage;
    use std::io::{BufRead, BufReader, Read, Write};
    use std::net::TcpListener;
    use std::thread;

    #[test]
    fn post_chat_sends_json_to_fake_upstream() {
        let listener = TcpListener::bind("127.0.0.1:0").expect("bind listener");
        let addr = listener.local_addr().expect("local addr");

        let server = thread::spawn(move || {
            let (mut stream, _) = listener.accept().expect("accept");
            let mut reader = BufReader::new(stream.try_clone().expect("clone stream"));
            let mut request_line = String::new();
            reader.read_line(&mut request_line).expect("request line");

            let mut content_length = 0usize;
            loop {
                let mut line = String::new();
                reader.read_line(&mut line).expect("header line");
                let trimmed = line.trim_end_matches(['\r', '\n']);
                if trimmed.is_empty() {
                    break;
                }
                if let Some((key, value)) = trimmed.split_once(':') {
                    if key.eq_ignore_ascii_case("Content-Length") {
                        content_length = value.trim().parse::<usize>().expect("content length");
                    }
                }
            }

            let mut body = vec![0u8; content_length];
            reader.read_exact(&mut body).expect("request body");
            let req = format!("{}\n{}", request_line, String::from_utf8_lossy(&body));
            assert!(req.starts_with("POST /api/chat HTTP/1.1"));
            assert!(req.contains("\"session_id\":7"));
            assert!(req.contains("\"id\":1"));

            let body = "event: done\ndata: done\n\n";
            let response = format!(
                "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream\r\nContent-Length: {}\r\n\r\n{}",
                body.len(),
                body
            );
            stream
                .write_all(response.as_bytes())
                .expect("write response");
        });

        let payload = PlatformChatRequest {
            session_id: 7,
            model: "gpt-5.4".to_string(),
            messages: vec![PlatformMessage {
                id: 1,
                role: "user".to_string(),
                content: "hello".to_string(),
            }],
        };

        let output = post_chat(&format!("http://{}", addr), "fake-token", &payload)
            .expect("request should succeed");
        assert!(output.contains("event: done"));

        server.join().expect("server join");
    }
}
