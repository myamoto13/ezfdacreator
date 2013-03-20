package com.extia.fdaprocessor.data;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author Michael Cortes
 *
 */
public class FicheAduction {
	
	private String identifiantSite;
	private String description;
	private String descDerivation;

	private List<Cable> cableList;
	private List<Jaretiere> jaretiereBPIList; //boitier Pied Immeuble
	private List<Jaretiere> jaretiereNROList;// Noeud Racordement Optique

	public FicheAduction() {
		cableList = new ArrayList<Cable>();
		jaretiereBPIList = new ArrayList<Jaretiere>();
		jaretiereNROList = new ArrayList<Jaretiere>();
	}

	public List<Cable> getCableList() {
		return cableList;
	}

	public void addCableList(Cable... cables) {
		if(cables != null){
			cableList.addAll(Arrays.asList(cables));
		}
	}

	public List<Jaretiere> getJaretiereBPIList() {
		return jaretiereBPIList;
	}

	public void addJaretiereBPIList(Jaretiere... jaretieres) {
		if(jaretieres != null){
			jaretiereBPIList.addAll(Arrays.asList(jaretieres));
		}
	}

	public List<Jaretiere> getJaretiereNROList() {
		return jaretiereNROList;
	}

	public void addJaretiereNROList(Jaretiere... jaretieres) {
		if(jaretieres != null){
			jaretiereNROList.addAll(Arrays.asList(jaretieres));
		}
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public String getDescDerivation() {
		return descDerivation;
	}

	public void setDescDerivation(String descDerivation) {
		this.descDerivation = descDerivation;
	}
	
	public String getIdentifiantSite() {
		return identifiantSite;
	}

	public void setIdentifiantSite(String identifiantSite) {
		this.identifiantSite = identifiantSite;
	}

	@Override
	public String toString() {
		return "FicheDAduction [identifiantSite=" + identifiantSite
				+ ", description=" + description + ", descDerivation="
				+ descDerivation + ", cableList=" + cableList
				+ ", jaretiereBPIList=" + jaretiereBPIList
				+ ", jaretiereNROList=" + jaretiereNROList + "]";
	}

}
