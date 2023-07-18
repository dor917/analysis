package com.daon.analysis.dto;

import java.util.Objects;

public class DuplicationData {
    private String pointName;
    private String treeName;
    private Double diameter;

    public String getPointName() {
        return pointName;
    }

    public void setPointName(String pointName) {
        this.pointName = pointName;
    }

    public String getTreeName() {
        return treeName;
    }

    public void setTreeName(String treeName) {
        this.treeName = treeName;
    }

    public Double getDiameter() {
        return diameter;
    }

    public void setDiameter(Double diameter) {
        this.diameter = diameter;
    }

    @Override
    public String toString() {
        return "DuplicationData{" +
                "pointName='" + pointName + '\'' +
                ", treeName='" + treeName + '\'' +
                ", diameter=" + diameter +
                '}';
    }


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        DuplicationData that = (DuplicationData) o;
        return Objects.equals(pointName, that.pointName) && Objects.equals(treeName, that.treeName);
    }

    @Override
    public int hashCode() {
        return Objects.hash(pointName, treeName);
    }
}
