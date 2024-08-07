import json
import asyncio
from requests import request
from os import listdir


class Gestao:
    def __init__(self) -> None:
        self.baseUrl = "http://gestao.faridnet.com.br/"
        self.token = None

    async def auth(self) -> None:
        try:
            response = request(
                method="post",
                url=f"{self.baseUrl}/Account/Login",
                headers={
                    "Content-Type": "application/json",
                },
                data=json.dumps({
                    'Username': "admin",
                    'Password': "ms@farid123"
                })
            )

            self.token = response.headers["Set-Cookie"]

        except Exception as error:
            print(error)

    async def createUsers(self, user_data: dict):
        try:
            response = request(
                method="post",
                url=f"{self.baseUrl}/services/Administration/User/Create",
                headers={
                    'Content-Type': 'application/json',
                    'Cookie': self.token
                },
                data=json.dumps({
                    "Entity": user_data
                })
            ).text

            return json.loads(response)

        except Exception as error:
            print(
                f"Ocorreu um problema. \n {error}")
