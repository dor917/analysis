package com.daon.analysis.dto;

import java.util.Objects;

public class DuplicationData {
    private String moutinaName;

    private String year;
    private String pointName;
    private String treeName;
    private Double diameter;
    private int cnt;

    public String getMoutinaName() {
        return moutinaName;
    }

    public void setMoutinaName(String moutinaName) {
        this.moutinaName = moutinaName;
    }

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

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public Double getDiameter() {
        return diameter;
    }

    public void setDiameter(Double diameter) {
        this.diameter = diameter;
    }


    public int getCnt() {
        return cnt;
    }

    public void setCnt(int cnt) {
        this.cnt = cnt;
    }

    @Override
    public String toString() {
        return "DuplicationData{" +
                "year='" + year + '\'' +
                "pointName='" + pointName + '\'' +
                ", treeName='" + treeName + '\'' +
                ", diameter=" + diameter +
                ", cnt=" + cnt +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        DuplicationData that = (DuplicationData) o;
        return Objects.equals(pointName, that.pointName) && Objects.equals(treeName, that.treeName) && Objects.equals(moutinaName, that.moutinaName) && Objects.equals(year, that.year);
    }

    @Override
    public int hashCode() {
        return Objects.hash(pointName, treeName);
    }
}
