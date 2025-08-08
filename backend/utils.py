import os
import json
import logging
import requests
import dataclasses

from typing import List, NamedTuple

DEBUG = os.environ.get("DEBUG", "false")
if DEBUG.lower() == "true":
    logging.basicConfig(level=logging.DEBUG)

AZURE_SEARCH_PERMITTED_GROUPS_COLUMN = os.environ.get(
    "AZURE_SEARCH_PERMITTED_GROUPS_COLUMN"
)


class JSONEncoder(json.JSONEncoder):
    def default(self, o):
        if dataclasses.is_dataclass(o):
            return dataclasses.asdict(o)
        return super().default(o)


async def format_as_ndjson(r):
    try:
        async for event in r:
            yield json.dumps(event, cls=JSONEncoder) + "\n"
    except Exception as error:
        logging.exception("Exception while generating response stream: %s", error)
        yield json.dumps({"error": str(error)})


def parse_multi_columns(columns: str) -> list:
    if "|" in columns:
        return columns.split("|")
    else:
        return columns.split(",")


def fetchUserGroups(userToken, nextLink=None):
    # Recursively fetch group membership
    if nextLink:
        endpoint = nextLink
    else:
        endpoint = "https://graph.microsoft.com/v1.0/me/transitiveMemberOf?$select=id"

    headers = {"Authorization": "bearer " + userToken}
    try:
        r = requests.get(endpoint, headers=headers)
        if r.status_code != 200:
            logging.error(f"Error fetching user groups: {r.status_code} {r.text}")
            return []

        r = r.json()
        if "@odata.nextLink" in r:
            nextLinkData = fetchUserGroups(userToken, r["@odata.nextLink"])
            r["value"].extend(nextLinkData)

        return r["value"]
    except Exception as e:
        logging.error(f"Exception in fetchUserGroups: {e}")
        return []


def generateFilterString(userToken):
    # Get list of groups user is a member of
    userGroups = fetchUserGroups(userToken)

    # Construct filter string
    if not userGroups:
        logging.debug("No user groups found")

    group_ids = ", ".join([obj["id"] for obj in userGroups])
    return f"{AZURE_SEARCH_PERMITTED_GROUPS_COLUMN}/any(g:search.in(g, '{group_ids}'))"


def format_non_streaming_response(chatCompletion, history_metadata, apim_request_id):
    response_obj = {
        "id": chatCompletion.id,
        "model": chatCompletion.model,
        "created": chatCompletion.created,
        "object": chatCompletion.object,
        "choices": [{"messages": []}],
        "history_metadata": history_metadata,
        "apim-request-id": apim_request_id,
    }

    if len(chatCompletion.choices) > 0:
        message = chatCompletion.choices[0].message
        if message:
            if hasattr(message, "context"):
                response_obj["choices"][0]["messages"].append(
                    {
                        "role": "tool",
                        "content": json.dumps(message.context),
                    }
                )
            response_obj["choices"][0]["messages"].append(
                {
                    "role": "assistant",
                    "content": message.content,
                }
            )
            return response_obj

    return {}

def format_stream_response(chatCompletionChunk, history_metadata, apim_request_id):
    response_obj = {
        "id": chatCompletionChunk.id,
        "model": chatCompletionChunk.model,
        "created": chatCompletionChunk.created,
        "object": chatCompletionChunk.object,
        "choices": [{"messages": []}],
        "history_metadata": history_metadata,
        "apim-request-id": apim_request_id,
    }

    if len(chatCompletionChunk.choices) > 0:
        delta = chatCompletionChunk.choices[0].delta
        if delta:
            if hasattr(delta, "context"):
                messageObj = {"role": "tool", "content": json.dumps(delta.context)}
                response_obj["choices"][0]["messages"].append(messageObj)
                return response_obj
            if delta.role == "assistant" and hasattr(delta, "context"):
                messageObj = {
                    "role": "assistant",
                    "context": delta.context,
                }
                response_obj["choices"][0]["messages"].append(messageObj)
                return response_obj
            if delta.tool_calls:
                messageObj = {
                    "role": "tool",
                    "tool_calls": {
                        "id": delta.tool_calls[0].id,
                        "function": {
                            "name" : delta.tool_calls[0].function.name,
                            "arguments": delta.tool_calls[0].function.arguments
                        },
                        "type": delta.tool_calls[0].type
                    }
                }
                if hasattr(delta, "context"):
                    messageObj["context"] = json.dumps(delta.context)
                response_obj["choices"][0]["messages"].append(messageObj)
                return response_obj
            else:
                if delta.content:
                    messageObj = {
                        "role": "assistant",
                        "content": delta.content,
                    }
                    response_obj["choices"][0]["messages"].append(messageObj)
                    return response_obj

    return {}


def format_pf_non_streaming_response(
    chatCompletion, history_metadata, response_field_name, citations_field_name, message_uuid=None
):
    if chatCompletion is None:
        logging.error(
            "chatCompletion object is None - Increase PROMPTFLOW_RESPONSE_TIMEOUT parameter"
        )
        return {
            "error": "No response received from promptflow endpoint increase PROMPTFLOW_RESPONSE_TIMEOUT parameter or check the promptflow endpoint."
        }
    if "error" in chatCompletion:
        logging.error(f"Error in promptflow response api: {chatCompletion['error']}")
        return {"error": chatCompletion["error"]}

    logging.debug(f"chatCompletion: {chatCompletion}")
    try:
        messages = []
        if response_field_name in chatCompletion:
            messages.append({
                "role": "assistant",
                "content": chatCompletion[response_field_name] 
            })
        if citations_field_name in chatCompletion:
            citation_content= {"citations": chatCompletion[citations_field_name]}
            messages.append({ 
                "role": "tool",
                "content": json.dumps(citation_content)
            })

        response_obj = {
            "id": chatCompletion["id"],
            "model": "",
            "created": "",
            "object": "",
            "history_metadata": history_metadata,
            "choices": [
                {
                    "messages": messages,
                }
            ]
        }
        return response_obj
    except Exception as e:
        logging.error(f"Exception in format_pf_non_streaming_response: {e}")
        return {}


def convert_to_pf_format(input_json, request_field_name, response_field_name):
    output_json = []
    logging.debug(f"Input json: {input_json}")
    # align the input json to the format expected by promptflow chat flow
    for message in input_json["messages"]:
        if message:
            if message["role"] == "user":
                new_obj = {
                    "inputs": {request_field_name: message["content"]},
                    "outputs": {response_field_name: ""},
                }
                output_json.append(new_obj)
            elif message["role"] == "assistant" and len(output_json) > 0:
                output_json[-1]["outputs"][response_field_name] = message["content"]
    logging.debug(f"PF formatted response: {output_json}")
    return output_json


def comma_separated_string_to_list(s: str) -> List[str]:
    '''
    Split comma-separated values into a list.
    '''
    return s.strip().replace(' ', '').split(',')


class ModelApiConfig(NamedTuple):
    """Configuration for different model API patterns"""
    uses_max_completion_tokens: bool
    uses_responses_endpoint: bool  # True for /responses, False for /deployments/{model}/chat/completions
    api_version: str
    

def get_model_api_config(model_name: str) -> ModelApiConfig:
    """
    Get the complete API configuration for a model.
    
    Args:
        model_name: The model name/deployment name
    
    Returns:
        ModelApiConfig with endpoint pattern, API version, and token parameter info
    """
    model_lower = model_name.lower()
    
    # GPT-5 series models use different endpoint and API version
    if any(pattern in model_lower for pattern in ['gpt-5', 'gpt5']):
        return ModelApiConfig(
            uses_max_completion_tokens=True,
            uses_responses_endpoint=True,
            api_version="2025-04-01-preview"
        )
    
    # o1 series models use max_completion_tokens but traditional endpoint
    if model_lower.startswith('o1') or 'o1-' in model_lower:
        return ModelApiConfig(
            uses_max_completion_tokens=True,
            uses_responses_endpoint=False,
            api_version="2024-05-01-preview"
        )
        
    # Default configuration for older models (GPT-4, GPT-3.5, etc.)
    return ModelApiConfig(
        uses_max_completion_tokens=False,
        uses_responses_endpoint=False,
        api_version="2024-05-01-preview"
    )


def uses_max_completion_tokens(model_name: str) -> bool:
    """
    Determine if a model uses max_completion_tokens instead of max_tokens.
    
    Args:
        model_name: The model name/deployment name
    
    Returns:
        True if model uses max_completion_tokens, False if it uses max_tokens
    """
    return get_model_api_config(model_name).uses_max_completion_tokens


def format_request_for_model(model_args: dict, model_name: str) -> dict:
    """
    Format request parameters based on the model's API requirements.
    
    GPT-5 models use the Responses API which requires different parameter names:
    - 'messages' becomes 'input'
    - Other parameters remain the same
    
    Args:
        model_args: Original request arguments
        model_name: The model name
        
    Returns:
        Formatted request arguments for the specific model
    """
    config = get_model_api_config(model_name)
    
    # For GPT-5 models using Responses API
    if config.uses_responses_endpoint:
        formatted_args = model_args.copy()
        
        # Transform 'messages' to 'input' for Responses API
        if 'messages' in formatted_args:
            messages = formatted_args.pop('messages')
            formatted_args['input'] = messages
        
        # Transform 'max_completion_tokens' to 'max_output_tokens' for Responses API
        if 'max_completion_tokens' in formatted_args:
            max_tokens = formatted_args.pop('max_completion_tokens')
            formatted_args['max_output_tokens'] = max_tokens
        
        # Handle extra_body parameter (move contents to top level for Responses API)
        if 'extra_body' in formatted_args:
            extra_body = formatted_args.pop('extra_body')
            # Move security context to top level if present
            if isinstance(extra_body, dict) and 'user_security_context' in extra_body:
                formatted_args['user_security_context'] = extra_body['user_security_context']
        
        # Remove parameters not supported by Responses API
        unsupported_params = ['stop', 'temperature', 'top_p']  # GPT-5 doesn't support these parameters
        for param in unsupported_params:
            if param in formatted_args:
                formatted_args.pop(param)
                
        return formatted_args
    
    # For traditional models, return unchanged
    return model_args

