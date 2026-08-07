use crate::engine::{get_token, TokenDef};
use crate::proxy::sse;
use crate::proxy::translate::{
    build_non_stream_response, build_stream_chunk, build_stream_stop_chunk, to_platform_request,
    OpenAiChatCompletionRequest,
};
use crate::proxy::upstream;
use std::env;
use std::io::{self, BufRead, BufReader, Read, Write};
use std::net::{TcpListener, TcpStream};

const DEFAULT_ADDR: &str = "127.0.0.1:8788";
const DEFAULT_BASE_URL: &str = "https://ai.solotopiax.com";
const DEFAULT_SESSION_ID: u64 = 1047;

pub fn serve(tokens: &[TokenDef]) -> io::Result<()> {
    let token = get_token(tokens, None).ok_or_else(|| {
        io::Error::new(
            io::ErrorKind::NotFound,
            "没找到 token。请先执行 lazytoken add 或配置 token 文件。",
        )
    })?;

    let addr = env::var("LAZYTOKEN_PROXY_ADDR").unwrap_or_else(|_| DEFAULT_ADDR.to_string());
    let listener = TcpListener::bind(&addr)?;
    println!("lazytoken proxy listening on http://{}", addr);

    for incoming in listener.incoming() {
        let mut stream = match incoming {
            Ok(s) => s,
            Err(e) => {
                eprintln!("accept 失败: {}", e);
                continue;
            }
        };

        if let Err(e) = handle_connection(&mut stream, token) {
            let status = if e.kind() == io::ErrorKind::InvalidInput {
                400
            } else {
                500
            };
            let _ = write_http_error(&mut stream, status, &e.to_string());
        }
    }

    Ok(())
}

fn handle_connection(stream: &mut TcpStream, token: &str) -> io::Result<()> {
    let request = read_http_request(stream)?;

    if request.method == "OPTIONS" {
        write_http_options(stream)?;
        return Ok(());
    }

    if request.method != "POST" || request.path != "/v1/chat/completions" {
        write_http_error(stream, 404, "仅支持 POST /v1/chat/completions")?;
        return Ok(());
    }

    let openai_req: OpenAiChatCompletionRequest =
        serde_json::from_slice(&request.body).map_err(|e| {
            io::Error::new(io::ErrorKind::InvalidInput, format!("JSON 解析失败: {}", e))
        })?;

    let session_id = resolve_session_id();
    let base_url = env::var("LAZYTOKEN_BASE_URL").unwrap_or_else(|_| DEFAULT_BASE_URL.to_string());
    let platform_req = to_platform_request(&openai_req, session_id)
        .map_err(|e| io::Error::new(io::ErrorKind::InvalidInput, e))?;

    let upstream_payload =
        upstream::post_chat(&base_url, token, &platform_req).map_err(io::Error::other)?;

    if openai_req.stream.unwrap_or(false) {
        write_openai_stream(stream, &openai_req.model, &upstream_payload)
    } else {
        write_openai_non_stream(stream, &openai_req.model, &upstream_payload)
    }
}

fn resolve_session_id() -> u64 {
    if let Ok(raw) = env::var("LAZYTOKEN_SESSION_ID") {
        if let Ok(id) = raw.parse::<u64>() {
            return id;
        }
    }
    DEFAULT_SESSION_ID
}

fn write_openai_non_stream(
    stream: &mut TcpStream,
    model: &str,
    upstream_sse: &str,
) -> io::Result<()> {
    let text = sse::extract_content_chunks(upstream_sse).join("");
    let body = build_non_stream_body(model, &text);
    write_http_json_ok(stream, &body)
}

fn write_openai_stream(stream: &mut TcpStream, model: &str, upstream_sse: &str) -> io::Result<()> {
    write!(
        stream,
        "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream; charset=utf-8\r\nCache-Control: no-cache\r\nConnection: close\r\nAccess-Control-Allow-Origin: *\r\n\r\n"
    )?;

    let body = build_stream_body(model, upstream_sse);
    stream.write_all(body.as_bytes())?;
    stream.flush()
}

fn build_non_stream_body(model: &str, content: &str) -> String {
    build_non_stream_response(model, content).to_string()
}

fn build_stream_body(model: &str, upstream_sse: &str) -> String {
    let chunk_id = format!("chatcmpl-{}", chrono::Utc::now().timestamp_millis());
    let mut out = String::new();
    for chunk in sse::extract_content_chunks(upstream_sse) {
        let payload = build_stream_chunk(model, &chunk_id, &chunk);
        out.push_str(&format!("data: {}\n\n", payload));
    }

    let stop_payload = build_stream_stop_chunk(model, &chunk_id);
    out.push_str(&format!("data: {}\n\n", stop_payload));
    out.push_str("data: [DONE]\n\n");
    out
}

fn write_http_json_ok(stream: &mut TcpStream, body: &str) -> io::Result<()> {
    write!(
        stream,
        "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: {}\r\nAccess-Control-Allow-Origin: *\r\n\r\n{}",
        body.len(),
        body
    )
}

fn write_http_options(stream: &mut TcpStream) -> io::Result<()> {
    write!(
        stream,
        "HTTP/1.1 204 No Content\r\nAccess-Control-Allow-Origin: *\r\nAccess-Control-Allow-Methods: POST, OPTIONS\r\nAccess-Control-Allow-Headers: Content-Type, Authorization\r\n\r\n"
    )
}

fn write_http_error(stream: &mut TcpStream, status: u16, message: &str) -> io::Result<()> {
    let body = serde_json::json!({
        "error": {
            "message": message,
            "type": "invalid_request_error",
        }
    })
    .to_string();
    write!(
        stream,
        "HTTP/1.1 {} ERROR\r\nContent-Type: application/json\r\nContent-Length: {}\r\nAccess-Control-Allow-Origin: *\r\n\r\n{}",
        status,
        body.len(),
        body
    )
}

struct HttpRequest {
    method: String,
    path: String,
    body: Vec<u8>,
}

fn read_http_request(stream: &mut TcpStream) -> io::Result<HttpRequest> {
    let mut reader = BufReader::new(stream);
    let mut request_line = String::new();
    reader.read_line(&mut request_line)?;
    if request_line.trim().is_empty() {
        return Err(io::Error::new(io::ErrorKind::UnexpectedEof, "空请求"));
    }

    let mut parts = request_line.split_whitespace();
    let method = parts.next().unwrap_or_default().to_string();
    let path = parts.next().unwrap_or_default().to_string();

    let mut content_length = 0usize;
    loop {
        let mut line = String::new();
        reader.read_line(&mut line)?;
        let line = line.trim_end_matches(['\r', '\n']);
        if line.is_empty() {
            break;
        }
        if let Some((key, value)) = line.split_once(':') {
            if key.eq_ignore_ascii_case("Content-Length") {
                content_length = value.trim().parse::<usize>().unwrap_or(0);
            }
        }
    }

    let mut body = vec![0u8; content_length];
    if content_length > 0 {
        reader.read_exact(&mut body)?;
    }

    Ok(HttpRequest { method, path, body })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_session_id_reads_env() {
        std::env::set_var("LAZYTOKEN_SESSION_ID", "7788");
        assert_eq!(resolve_session_id(), 7788);
        std::env::remove_var("LAZYTOKEN_SESSION_ID");
    }

    #[test]
    fn build_stream_body_emits_done_marker() {
        let upstream =
            "event: content\ndata: 你\n\nevent: content\ndata: 好\n\nevent: done\ndata: done\n\n";
        let body = build_stream_body("gpt-5.4", upstream);
        assert!(body.contains("chat.completion.chunk"));
        assert!(body.contains("data: [DONE]"));
    }
}
