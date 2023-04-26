import openai

openai.api_key = "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"


def completion(prompt):
    completions = openai.Completion.create(
        engine="text-davinci-002",
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.8,
    )

    message = completions.choices[0].text
    return message


if __name__ == '__main__':
    try:
        print(completion("test"))
        print(completion("stm32f103vct6串口1初始化代码"))
    except Exception as e:
        print(print(e))