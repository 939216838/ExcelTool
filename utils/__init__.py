def convert_to_snake_case(s):
    snake_case = ''
    for i, c in enumerate(s):
        if c.isupper() and i > 0:
            snake_case += '_'
        snake_case += c.lower()
    return snake_case

if __name__ == '__main__':
    print(convert_to_snake_case("fullOnlineNonNaturalPersonList"))