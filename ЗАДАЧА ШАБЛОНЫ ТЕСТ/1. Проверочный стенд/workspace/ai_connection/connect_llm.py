from __future__ import annotations

from abc import ABC, abstractmethod
import ssl
import httpx

from workspace.workspace_params.modules import create_request_ssl, le
from workspace.workspace_params.settings import (
    GIGACHAT_CERT, GIGACHAT_KEY
)

class AsyncHttpRequest(ABC):
    """
    Класс для выполнения асинхронных HTTP-запросов.
    Этот класс реализует метод `make_request` для выполнения асинхронных 
    HTTP-запросов и обработки различных типов данных и методов.
    """
    async def make_request(
        self, url, method="GET", headers=None, data=None, files=None, 
        cert=None, key=None, verify=True, cacert=None, to_json=True,
        timeout=900.0
    ):

        context = await create_request_ssl(cert, key, False)

        if method.upper() == "GET":
            async with httpx.AsyncClient(
                verify=context, timeout=timeout
            ) as client:
                response = await client.get(url, headers=headers, params=data)
        elif method.upper() == "POST":
            if files:
                async with httpx.AsyncClient(
                    verify=context, timeout=timeout
                ) as client:
                    response = await client.post(
                        url, headers=headers, data=data, files=files
                    )
            else:
                async with httpx.AsyncClient(
                    verify=context, timeout=timeout
                ) as client:
                    response = await client.post(
                        url, headers=headers, json=data
                    )
        elif method.upper() == "PUT":
            if files:
                async with httpx.AsyncClient(
                    verify=context, timeout=timeout
                ) as client:
                    response = await client.put(
                        url, headers=headers, data=data, files=files
                    )
            else:
                async with httpx.AsyncClient(
                    verify=context, timeout=timeout
                ) as client:
                    response = await client.put(
                        url, headers=headers, json=data
                    )
        elif method.upper() == "DELETE":
            async with httpx.AsyncClient(
                verify=context, timeout=timeout
            ) as client:
                response = await client.delete(
                    url, headers=headers, params=data
                )
        else:
            raise ValueError(f"Неподдерживаемый тип запроса: {method}")
        try:
            response.raise_for_status()
            return response.json() if to_json else response
        except Exception as error:
            le(f"Ошибка запроса: {error}.\nОтвет: {response.text}")


class AbstractRequest(ABC):
    """
    Абстрактный базовый класс для выполнения HTTP-запросов.
    Этот класс определяет интерфейс c абстрактным методом `make_request`, который должен быть реализован любым подклассом.
    """

    @abstractmethod
    async def make_request(self, url, method="GET", headers=None, data=None, files=None, cert=None, key=None, verify=True,
                           cacert=None, to_json=True):
        """
        Абстрактный метод для выполнения HTTP-запроса.

        Параметры:
            url (str): URL, по которому будет отправлен запрос.
            method (str, optional): HTTP-метод, который будет использоваться (по умолчанию "GET").
            headers (dict, optional): Заголовки, которые будут включены в запрос (по умолчанию None).
            data (dict, optional): Данные, которые будут отправлены в теле запроса (по умолчанию None).
            files (dict, optional): Файлы, которые будут отправлены в теле запроса (по умолчанию None).
            cert (str or tuple, optional): Путь к сертификату клиента или кортеж (путь к сертификату, путь к ключу) (по умолчанию None).
            key (str, optional): Путь к ключу клиента (по умолчанию None).
            verify (bool or str, optional): Указывает, следует ли проверять SSL-сертификат сервера (по умолчанию True).
                                           Если указано путь к файлю, он будет использован как CA_BUNDLE.
            cacert (str, optional): Путь к CA_BUNDLE (по умолчанию None).

        Возвращает:
            dict: JSON-ответ от сервера.

        Исключения:
            NotImplementedError: Если метод не реализован в подкласса.
        """
        pass


class SyncHttpRequest(AbstractRequest):
    """
    Класс для выполнения синхронных HTTP-запросов.
    Этот класс реализует метод `make_request` для выполнения синхронных HTTP-запросов и обработки различных типов данных и методов.
    """
    def make_request(self, url, method="GET", headers=None, data=None, 
        files=None, cert=GIGACHAT_CERT, key=GIGACHAT_KEY, verify=True,
        cacert=None, to_json=True, timeout=900.0
    ):

        if cacert:
            verify = cacert

        context = self.sync_create_request_ssl(cert, key, False)

        # Выбор метода запроса
        if method.upper() == "POST":
            if files:
                with httpx.Client(verify=context, timeout=timeout) as client:
                    response = client.post(
                        url, headers=headers, data=data, files=files
                    )
            else:
                with httpx.Client(verify=context, timeout=timeout) as client:
                    response = client.post(
                        url, headers=headers, json=data
                    )
        else:
            raise ValueError(f"Неподдерживаемый тип запроса: {method}")
        try:
            response.raise_for_status()
            return response.json() if to_json else response
        except Exception as error:
            le(f"Ошибка запроса: {error}.\nОтвет: {response.text}")

    @staticmethod
    def sync_create_request_ssl(cert,key, verify):
        ssl_context = ssl.create_default_context()
        ssl_context.load_cert_chain(cert, key)
        ssl_context.check_hostname = verify
        ssl_context.verify_mode = ssl.CERT_NONE
        return ssl_context
