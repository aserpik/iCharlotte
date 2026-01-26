import requests
import datetime
from PySide6.QtCore import QThread, Signal
from .config import API_KEYS
from .utils import log_event

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

class LLMHandler:
    @staticmethod
    def create_cache(provider, model, content, system_instruction=None, ttl_minutes=60):
        """Creates a cache for Gemini and returns the cache name."""
        if provider != "Gemini":
            return None
            
        if not genai:
            raise ImportError("google.genai not installed")
            
        api_key = API_KEYS.get(provider)
        client = genai.Client(api_key=api_key)
        
        try:
            # Create config for cache
            cache_config = {
                "contents": [content],
                "ttl": f"{ttl_minutes*60}s"
            }
            if system_instruction:
                cache_config["system_instruction"] = system_instruction

            cache = client.caches.create(
                model=model,
                config=types.CreateCachedContentConfig(**cache_config)
            )
            return cache.name
        except Exception as e:
            print(f"Cache creation failed: {e}")
            return None

    @staticmethod
    def generate(provider, model, system_prompt, user_prompt, file_contents, settings, history=None):
        """
        Generic generator.
        If settings['stream'] is True, yields chunks (str).
        Otherwise returns full text (str).
        
        history: Optional list of dicts [{'role': 'user'|'assistant', 'content': str}, ...]
        """
        temp = settings.get('temperature', 1.0)
        top_p = settings.get('top_p', 0.95)
        max_tokens = settings.get('max_tokens', -1)
        thinking_level = settings.get('thinking_level', "None")
        do_stream = settings.get('stream', False)
        cache_name = settings.get('cache_name', None)
        
        api_key = API_KEYS.get(provider)
        if not api_key:
            raise ValueError(f"API Key for {provider} not found.")

        if provider == "Gemini":
            if not genai:
                raise ImportError("google.genai not installed")

            client = genai.Client(api_key=api_key)
            api_level = thinking_level.upper()
            
            config_params = {
                "temperature": temp,
                "top_p": top_p,
            }
            
            # Only add system_instruction if NOT using cache (cache has it)
            if not cache_name:
                config_params["system_instruction"] = system_prompt

            if max_tokens > 0:
                config_params["max_output_tokens"] = max_tokens
            if thinking_level != "None" and ("gemini-3" in model.lower() or "thinking" in model.lower()):
                config_params["thinking_config"] = {
                    "include_thoughts": True,
                    "thinking_level": api_level
                }
            if cache_name:
                    config_params["cached_content"] = cache_name

            # Prompt construction
            if history:
                # Convert generic history to Gemini contents
                gemini_contents = []
                for h in history:
                    role = "user" if h['role'] == "user" else "model"
                    gemini_contents.append(types.Content(role=role, parts=[types.Part(text=h['content'])]))
                
                # Current message
                full_prompt = user_prompt
                if file_contents and not cache_name:
                    full_prompt += "\n\n[ATTACHED FILES]:\n" + file_contents
                
                gemini_contents.append(types.Content(role="user", parts=[types.Part(text=full_prompt)]))
                
                final_contents = gemini_contents
            else:
                # Legacy/Simple mode
                full_prompt = user_prompt
                if file_contents and not cache_name:
                    # Only append files if NOT using cache (cache presumably has them)
                    full_prompt += "\n\n[ATTACHED FILES]:\n" + file_contents
                final_contents = full_prompt
                
            # Helper to run generation
            def _run_gen(cfg):
                if do_stream:
                    return client.models.generate_content_stream(
                        model=model,
                        contents=final_contents,
                        config=types.GenerateContentConfig(**cfg)
                    )
                else:
                    return client.models.generate_content(
                        model=model,
                        contents=final_contents,
                        config=types.GenerateContentConfig(**cfg)
                    )

            try:
                log_event(f"Sending request to Gemini (Model: {model})")
                response = _run_gen(config_params)
                log_event("Received response from Gemini")
            except Exception as e:
                log_event(f"Initial Gemini request failed: {e}", "warning")
                # Fallback retry without thinking if failed
                if "thinking_config" in config_params:
                    log_event("Retrying without thinking_config...", "warning")
                    del config_params["thinking_config"]
                    try:
                        response = _run_gen(config_params)
                    except Exception as e2:
                        log_event(f"Fallback request failed: {e2}", "error")
                        raise
                else:
                    raise

            if do_stream:
                # We must keep client reference alive for the generator to work
                def stream_wrapper(client_ref, stream_resp):
                    try:
                        for chunk in stream_resp:
                            yield chunk.text
                    except Exception as e:
                        log_event(f"Stream error: {e}", "error")
                        yield f"[Stream Error: {e}]"
                return stream_wrapper(client, response)
            else:
                try:
                    # Check for safety blocks or empty responses
                    if not response.candidates:
                            log_event("Gemini returned no candidates (blocked?)", "error")
                            return "[Error: The AI response was blocked or empty.]"
                    
                    if hasattr(response, 'text'):
                            return response.text
                    else:
                            return response.candidates[0].content.parts[0].text
                except Exception as e:
                    log_event(f"Error accessing response text: {e}", "error")
                    try:
                        log_event(f"Response dump: {response}", "info")
                    except: pass
                    raise

        elif provider == "OpenAI":
            # (Existing OpenAI logic - simplified: no streaming support added yet)
            url = "https://api.openai.com/v1/chat/completions"
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
            
            messages = [{"role": "system", "content": system_prompt}]
            if history:
                messages.extend(history)
            
            full_user_content = user_prompt
            if file_contents: full_user_content += "\n\n[ATTACHED FILES]:\n" + file_contents
            messages.append({"role": "user", "content": full_user_content})
            
            payload = {
                "model": model, "messages": messages,
                "temperature": temp, "top_p": top_p
            }
            if max_tokens > 0: payload["max_tokens"] = max_tokens
            
            resp = requests.post(url, headers=headers, json=payload)
            if resp.status_code == 200:
                return resp.json()['choices'][0]['message']['content']
            raise Exception(f"OpenAI Error {resp.status_code}: {resp.text}")

        elif provider == "Claude":
             # (Existing Claude logic - simplified)
            url = "https://api.anthropic.com/v1/messages"
            headers = {"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"}
            
            messages = []
            if history:
                messages.extend(history)
                
            full_user_content = user_prompt
            if file_contents: full_user_content += "\n\n[ATTACHED FILES]:\n" + file_contents
            messages.append({"role": "user", "content": full_user_content})
            
            claude_max = max_tokens if max_tokens > 0 else 8192
            payload = {
                "model": model, "max_tokens": claude_max, "temperature": temp, "top_p": top_p,
                "system": system_prompt, "messages": messages
            }
            resp = requests.post(url, headers=headers, json=payload)
            if resp.status_code == 200:
                return resp.json()['content'][0]['text']
            raise Exception(f"Claude Error {resp.status_code}: {resp.text}")
        
        return "Provider not implemented."

class LLMWorker(QThread):
    finished = Signal(str)
    new_token = Signal(str) # For streaming
    error = Signal(str)
    
    def __init__(self, provider, model, system, user, files, settings, history=None):
        super().__init__()
        self.provider = provider
        self.model = model
        self.system = system
        self.user = user
        self.files = files
        self.settings = settings
        self.history = history
        
    def run(self):
        try:
            result = LLMHandler.generate(
                self.provider, self.model, self.system, 
                self.user, self.files, self.settings,
                history=self.history
            )
            
            if self.settings.get('stream', False):
                full_text = ""
                for chunk in result:
                    if chunk:
                        full_text += chunk
                        self.new_token.emit(chunk)
                self.finished.emit(full_text)
            else:
                self.finished.emit(result)
                
        except Exception as e:
            self.error.emit(str(e))

class ModelFetcher(QThread):
    finished = Signal(str, list)
    error = Signal(str)

    def __init__(self, provider, api_key):
        super().__init__()
        self.provider = provider
        self.api_key = api_key

    def run(self):
        try:
            models = []
            if self.provider == "OpenAI":
                url = "https://api.openai.com/v1/models"
                headers = {"Authorization": f"Bearer {self.api_key}"}
                resp = requests.get(url, headers=headers)
                if resp.status_code == 200:
                    all_models = resp.json()['data']
                    models = [m['id'] for m in all_models if m['id'].startswith(('gpt', 'o1', 'o3'))]
                    models.sort(reverse=True)
                else:
                    raise Exception(f"OpenAI Error: {resp.status_code} {resp.text}")

            elif self.provider == "Gemini":
                if genai:
                    client = genai.Client(api_key=self.api_key)
                    # Try iterating; if fails, catch exception
                    found_models = list(client.models.list())
                    for m in found_models:
                        name = m.name if hasattr(m, 'name') else str(m)
                        if name.startswith('models/'):
                            name = name[7:]
                        if 'gemini' in name.lower():
                            models.append(name)
                else:
                    raise ImportError("google.genai not installed")
                        
                models.sort(reverse=True)

            elif self.provider == "Claude":
                url = "https://api.anthropic.com/v1/models"
                headers = {
                    "x-api-key": self.api_key,
                    "anthropic-version": "2023-06-01"
                }
                resp = requests.get(url, headers=headers)
                if resp.status_code == 200:
                    data = resp.json()
                    models = [m['id'] for m in data['data']]
                    models.sort(reverse=True)
                else:
                    raise Exception(f"Claude Error: {resp.status_code} {resp.text}")
            
            self.finished.emit(self.provider, models)

        except Exception as e:
            self.error.emit(str(e))