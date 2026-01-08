# llm_client.py
import os
from loguru import logger
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

UPSTAGE_API_KEY = os.environ.get("UPSTAGE_API_KEY")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
GEMINI_API_KEY = os.environ.get("GEMINI")


def get_api_key_and_provider(model: str):
    """
    모델 이름으로부터 어떤 provider를 쓸지 + 어떤 API key를 쓸지 결정
    """
    provider = None
    api_key = None

    # OpenAI
    if model.startswith("gpt-") or model in ["gpt-4o", "gpt-4.1", "gpt-4o-mini"]:
        provider = "openai"
        api_key = OPENAI_API_KEY

    # Anthropic
    elif model.startswith("claude-"):
        provider = "anthropic"
        api_key = ANTHROPIC_API_KEY

    # Gemini
    elif model.startswith("gemini-"):
        provider = "gemini"
        api_key = GEMINI_API_KEY

    # Upstage(Solar)
    elif model.startswith("solar-"):
        provider = "upstage"
        api_key = UPSTAGE_API_KEY

    else:
        raise ValueError(f"지원되지 않는 모델: {model}")

    if api_key is None:
        raise ValueError(f"{provider}용 API 키가 설정되어 있지 않습니다. (model={model})")

    return provider, api_key


def get_client_for_model(model: str):
    """
    모델에 맞는 OpenAI 스타일 클라이언트 반환
    """
    provider, api_key = get_api_key_and_provider(model)

    if provider == "openai":
        client = OpenAI(api_key=api_key)
        
    elif provider == "upstage":
        client = OpenAI(api_key=api_key, base_url="https://api.upstage.ai/v1")

    elif provider == "anthropic":
        # TODO: 나중에 Anthropic SDK로 구현
        ...
    
    elif provider == "gemini":
        # TODO: 나중에 Gemini API로 구현
        ...
    
    else:
        raise ValueError(f"알 수 없는 provider: {provider}")

    return client, provider


def call_llm(
    model: str,
    messages,
    tools=None,
    tool_choice=None,
    **kwargs,
):
    """
    공통 chat completion 호출 래퍼
    - Agent에서 이 함수만 사용하게 해서, 모델/키/클라이언트 관리를 이 파일로 몰아줌
    """
    client, provider = get_client_for_model(model)

    payload = {
        "model": model,
        "messages": messages,
    }

    if tools is not None:
        payload["tools"] = tools
    if tool_choice is not None:
        payload["tool_choice"] = tool_choice

    # 기타 추가 kwargs (temperature, max_tokens 등)
    payload.update(kwargs)

    return client.chat.completions.create(**payload)

