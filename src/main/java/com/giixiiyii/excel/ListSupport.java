package com.giixiiyii.excel;

import java.util.ArrayList;
import java.util.Collection;

public abstract class ListSupport<T> extends ArrayList<T> {
    public ListSupport<T> append(T e) {
        add(e);
        return this;
    }

    public ListSupport<T> append(int i, T e) {
        add(i, e);
        return this;
    }

    public ListSupport<T> appendAll(Collection<? extends T> c) {
        if (c != null) addAll(c);
        return this;
    }
}