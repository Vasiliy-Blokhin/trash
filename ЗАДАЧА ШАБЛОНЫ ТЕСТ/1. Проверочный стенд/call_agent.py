import asyncio
from asgiref.sync import sync_to_async

from agent_check_document import AgentCheckDocument
from workspace.workspace_params.settings import REQUEST_PARAMS
from workspace.workspace_params.modules import li


async def call_agent(
    content,
    file_content,
    full_url,
    short_url,
    task_id,
):
    return await AgentCheckDocument(
        content=content['messages'][-1]['content'],
        request_data=content,
        file_content=file_content,
        full_url=full_url,
        short_url=short_url,
        task_id=task_id,
    )()


if __name__ == "__main__":
    li(asyncio.run(
        call_agent(
            content=REQUEST_PARAMS["request_data"],
            file_content=REQUEST_PARAMS["file_content"],
            full_url=REQUEST_PARAMS["full_url"],
            short_url=REQUEST_PARAMS["short_url"],
            task_id=REQUEST_PARAMS["task_id"]
        )
    ))
