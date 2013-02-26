package com.extia.fdaprocessor.data;

public class Jaretiere {

	private String tenant;
	private String aboutissant;
	private String ref;
	private String etat;
	private String commentaires;
	
	public String getTenant() {
		return tenant;
	}
	public void setTenant(String tenant) {
		this.tenant = tenant;
	}
	public String getAboutissant() {
		return aboutissant;
	}
	public void setAboutissant(String aboutissant) {
		this.aboutissant = aboutissant;
	}
	public String getRef() {
		return ref;
	}
	public void setRef(String ref) {
		this.ref = ref;
	}
	public String getEtat() {
		return etat;
	}
	public void setEtat(String etat) {
		this.etat = etat;
	}
	public String getCommentaires() {
		return commentaires;
	}
	public void setCommentaires(String commentaires) {
		this.commentaires = commentaires;
	}
	@Override
	public String toString() {
		return "Jaretiere [tenant=" + tenant + ", aboutissant=" + aboutissant
				+ ", ref=" + ref + ", etat=" + etat + ", commentaires="
				+ commentaires + "]";
	}
	
	
	
}
