# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import openai

openai.api_key = "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"


def completion(prompt):
    completions = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.8,
    )

    message = completions.choices[0].text
    return message

if __name__ == '__main__':
    completion("鸟")

print(completion("stm32f103vct6串口1初始化代码"))
