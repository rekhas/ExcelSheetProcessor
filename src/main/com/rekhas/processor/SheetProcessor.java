package main.com.rekhas.processor;

public abstract class SheetProcessor<T> {

    public abstract void process(long rowIndex, long columnIndex, String value);

}
