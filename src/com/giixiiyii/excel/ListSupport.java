package com.giixiiyii.excel;

import java.util.ArrayList;
import java.util.List;

public abstract class ListSupport<T> {
    List<T> data = new ArrayList();

    public int size() {
        return data.size();
    }

    public T get(int i) {
        return data.get(i);
    }

    public List<T> getData() {
        return data;
    }

    public ListSupport<T> add(T e) {
        data.add(e);
        return this;
    }

    public ListSupport<T> add(int i, T e) {
        data.add(i, e);
        return this;
    }

    public ListSupport<T> addAll(List<T> li) {
        if (li != null) data.addAll(li);
        return this;
    }
}