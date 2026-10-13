from workspace.ai_connection.ai_worker import AIWorker

class AgentAIWorker(AIWorker):
    def __init__(self, system_prompt='', content=''):
        super().__init__(system_prompt, content)