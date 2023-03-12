

if __name__ == '__main__':

    my_dict = {"1":None,"2":"","3":"","4":"","5":""}
    # ```python
    len_dict = len([k for k in my_dict if my_dict[k] is not None])
    print(len_dict)
    # ```

    # ```python
    len_dict = len([k for k, v in my_dict.items() if v is not None])
    print(len_dict)
    # ```
